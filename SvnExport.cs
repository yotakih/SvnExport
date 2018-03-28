using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using SharpSvn;
using System.Net;
using System.IO;

namespace SvnExportor
{
  class SvnExport
  {
    private SvnClient _svncl;
    private string _src;
    private Uri _uri;
    private Uri _reporoot;
    private string _repoIdntfr;
    private string[] _filterList;
    private ExcelStream _exlStm;
    private Dictionary<string, SvnChangeAction[]> _dicPths =
      new Dictionary<string, SvnChangeAction[]>();
    private long[] _revLst;
    
    public SvnExport(string src,string usr,string psswd)
    {
      this._src = src;
      _svncl = new SvnClient();
      _svncl.Authentication.DefaultCredentials = new NetworkCredential(usr, psswd);

      var tgt = new SvnPathTarget(this._src);
      SvnInfoEventArgs rslt;
      this._svncl.GetInfo(tgt, out rslt);
      this._uri = rslt.Uri;
      this._reporoot = rslt.RepositoryRoot;
      this._repoIdntfr =
        rslt.Uri.ToString().Replace(
          this._reporoot.ToString(), @""
        );
    }

    public void clearAuth()
    {
      //_svncl.Authentication.Clear();
      _svncl.Authentication.ClearAuthenticationCache();
    }

    public void readFilterFile(string filePath)
    {
      if (!File.Exists(filePath))
        throw new FileNotFoundException(string.Format("フィルター設定ファイルがありません\n{0}", filePath));
      string txt;
      using (var sr = new StreamReader(filePath, Encoding.GetEncoding(932)))
        txt = sr.ReadToEnd();
      this._filterList = txt.Replace("\r\n", "\n").Split('\n').Where(
        ln => ln.Trim().Length > 0 && ln.Trim().Substring(0, 1) != @"#").Select(ln => ln.ToLower()).ToArray();
    }

    public void export(string strRevLst,string toRootPath)
    {
#if DEBUG

#else
      try
      {
#endif
      if (Directory.Exists(toRootPath))
        throw new Exception(
            string.Format(@"すでに出力先フォルダが存在します。：{0}", toRootPath)
          );

      try
      {
        this._revLst = strRevLst.Split(',').Select(v => long.Parse(v.Trim())).ToArray();
      }catch(Exception e) {
        throw new Exception(@"抽出対象リビジョンリストの様式が不正です。確認してください。");
      }

      var expDir = Path.Combine(toRootPath, @"exportFiles");
      if (!Directory.Exists(toRootPath)) Directory.CreateDirectory(toRootPath);
      if (!Directory.Exists(expDir)) Directory.CreateDirectory(expDir);

      this._exlStm = new ExcelStream();
      this._exlStm.opnWrkBook(Path.Combine(toRootPath, @"対象一覧.xlsx"));

      ana_logs(this._revLst, toRootPath, expDir);

#if DEBUG
      this._exlStm.save();
      this._exlStm.close();
#else
      }
      catch(Exception e)
      {
        Console.WriteLine(e.Message);
      }
      finally
      {
        if (this._exlStm != null)
        {
          this._exlStm.save();
          this._exlStm.close();
        }
      }
#endif
    }

    private void ana_logs(long[] revLst,string toRootPath,string expDir)
    {
      for (var i = 0; i < revLst.Count(); i++)
        ana_log(revLst[i], expDir, i);
      this._exlStm.addSmrySht(this._revLst, this._dicPths);
    }

    private void ana_log(long rev,string expDir,int revIdx)
    {
      Console.WriteLine(string.Format(
        @"リビジョン：{0}　を抽出します。", rev
      ));
      var arg = new SvnLogArgs()
      {
        Range = new SvnRevisionRange(rev, rev)
      };
      Collection<SvnLogEventArgs> log;
      this._svncl.GetLog(this._uri, arg, out log);
      if (log.Count == 0)
        throw new Exception(string.Format(@"リビジョン:{0}のログがありません。処理を途中で終了します。", rev));
      ana_revFls(log[0],expDir,revIdx);
    }

    private void ana_revFls(SvnLogEventArgs logargs, string expDir,int revIdx)
    {
      var expArg = new SvnExportArgs()
      {
        Depth = SvnDepth.Files ,
        Overwrite = true
      };
      var revname = retPadLeftZero(logargs.Revision);
      this._exlStm.addRevSht(revname, logargs.LogMessage);
      
      foreach(var itm in logargs.ChangedPaths)
      {
        if (itm.NodeKind != SvnNodeKind.File) continue;
        var path = itm.Path.Replace(@"/", @"\");
        var hitFlg = false;
        foreach (var flt in this._filterList)
          if (itm.Path.ToLower().IndexOf(flt) > 0)
          {
            hitFlg = true;
            break;
          }
        if (hitFlg) addDicPaths(path, revIdx, itm.Action);
        var trgtStr = @"";
        if (hitFlg && itm.Action != SvnChangeAction.Delete)
        {
          
          var srcPth = Path.Combine(this._reporoot.ToString(),
            itm.RepositoryPath.ToString());
          var dstPth =
            expDir + itm.Path.Replace(this._repoIdntfr, @"").Replace(@"/", @"\");
          if (!Directory.Exists(Path.GetDirectoryName(dstPth)))
            Directory.CreateDirectory(Path.GetDirectoryName(dstPth));

          this._svncl.Export(new SvnUriTarget(srcPth, logargs.Revision), dstPth, expArg);

          trgtStr = @"○";
        }

        this._exlStm.wrtRevRow(new string[]
          {
            itm.Action.ToString(),
            Path.GetDirectoryName(path),
            Path.GetFileName(path),
            trgtStr
          });
      }
    }

    private string retPadLeftZero(long l)
    {
      return l.ToString().PadLeft(5, '0');
    }

    private void addDicPaths(string pth, int idx,SvnChangeAction act)
    {
      if (!this._dicPths.ContainsKey(pth))
        this._dicPths.Add(pth, new SvnChangeAction[this._revLst.Count()]);
      this._dicPths[pth][idx] = act;
    }
  }
}
