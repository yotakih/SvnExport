using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using SharpSvn;
using AnalyzeCommandLineArgs;

namespace SvnExportor
{
  class Program
  {
    static string _src;
    static string _user;
    static string _psswd;
    static string _filerFl;
    static string _strRevList;
    static string _workdir;

    static CommandLineAnalyzer cmdAnz;

    static void Main(string[] args)
    {
      Console.WriteLine(@"抽出処理を開始します。");
      Console.WriteLine(@"");
#if DEBUG
      createAutoBatTest();
#else
      try 
      {
        setupCommandLineAnalizer();
        createAutoBat();
#endif
#if DEBUG

#else
    }
      catch (SvnClientUnrelatedResourcesException e)
      {
        Console.WriteLine(string.Format(@"エラー内容：{0}", @"指定されたリビジョンにこのリポジトリのログがありません"));
        Console.WriteLine(string.Format(@"エラークラス名：{0}", e.GetType().FullName));
        Console.WriteLine(string.Format(@"Message       ：{0}", e.Message));
      }
      catch (Exception e)
      {
        if (e.Message.Trim().Length > 0)
        {
          Console.WriteLine(string.Format(@"エラークラス名：{0}", e.GetType().FullName));
          Console.WriteLine(string.Format(@"エラー内容    ：{0}", e.Message));
        }
      }
#endif
#if DEBUG
      Console.WriteLine(@"");
      Console.WriteLine(@"出力処理が完了しました。");
      Console.WriteLine(@"終了するにはなにかキーを押してください。");
      Console.ReadKey();
#endif
    }

    static void setupCommandLineAnalizer()
    {
      cmdAnz = new CommandLineAnalyzer();
      CommandOption co_src = new CommandOption(@"-src", true, true
                                 , @"-src      [作業コピーフォルダ名]       コピー元リポジトリ作業コピーパスを指定"
                                 , (arg) => { if (!Directory.Exists(arg)) throw new Exception(@"-srcフォルダがありません"); });
      CommandOption co_user = new CommandOption(@"-user", true, true
                                 , @"-user     [-srcリポジトリのユーザ]     ログイン用ユーザ名を指定"
                                 , null);
      CommandOption co_psswd = new CommandOption(@"-psswd", true, true
                                 , @"-passwd   [-srcリポジトリのパスワード] ログイン用パスワードを指定"
                                 , null);
      CommandOption co_filerFl = new CommandOption(@"-filterFl", true, true
                                 , @"-filterFl [フィルタファイルパス]       抽出ファイルを絞り込むためのパターンを記載したファイルパス"
                                 , (arg) => { if (!File.Exists(arg)) throw new Exception(@"-filterFlファイルがありません"); });
      CommandOption co_strRevLst = new CommandOption(@"-strRevLst", true, true
                                 , @"-strRevLst  [抽出リビジョンリスト]     抽出対象リビジョン番号をカンマ区切りで指定"
                                 , (arg) => { long.Parse(arg); });
      CommandOption co_workDir = new CommandOption(@"-workDir", true, true
                                 , @"-workDir  [作業フォルダ名]             作業フォルダパスをフルパスで指定"
                                 , (arg) => {
                                   if (arg.IndexOf(@":") < 0) throw new Exception(@"-workDirはフルパスで指定します");
                                   if (Directory.Exists(arg)) throw new Exception(@"-workDirはすでに存在します。存在しないフォルダを指定します。処理中に自動作成します");
                                 });
      CommandOption co_Q = new CommandOption(@"/?", false, false
                                 , @"/?                                     ヘルプを表示"
                                 , null);
      cmdAnz.helpOption = co_Q;
      cmdAnz.addCommandOption(co_src);
      cmdAnz.addCommandOption(co_user);
      cmdAnz.addCommandOption(co_psswd);
      cmdAnz.addCommandOption(co_filerFl);
      cmdAnz.addCommandOption(co_strRevLst);
      cmdAnz.addCommandOption(co_workDir);

      if (!cmdAnz.analyze()) throw new Exception(@"");

      _src = co_src.arg;
      _user = co_user.arg;
      _psswd = co_psswd.arg;
      _filerFl = co_filerFl.arg;
      _strRevList = co_strRevLst.arg;
      _workdir = co_workDir.arg;
    }

    static void createAutoBat()
    {
      var se = new SvnExport(_src, _user, _psswd);
      se.readFilterFile(_filerFl);
      se.export(_strRevList, _workdir);
      
      se.clearAuth();
    }
#if DEBUG
    static void createAutoBatTest()
    {
      _src = @"C:\Temp\svn\testSrc\srcDir";
      _user = @"user";
      _psswd = @"pass";
      _filerFl = @"C:\Temp\svn\filter.txt";
      _strRevList = @"7,8,9,12,14,15,23,30,31,32";
      _workdir = @"C:\Temp\svn\exportPath\20180328";

      var se = new SvnExport(_src, _user, _psswd);
      se.readFilterFile(_filerFl);
      se.export(_strRevList, _workdir);

      se.clearAuth();
    }
#endif

  }
}
