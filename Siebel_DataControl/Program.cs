using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using SiebelBusObjectInterfaces;


namespace Siebel_DataControl
{
    class Program
    {
        enum ConnectMode { Local, Server};

        private static void printUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  Siebel_DataControl connectString");
            Console.WriteLine("    connectString\t - ServerConnectString | LocalConnectString");
            Console.WriteLine("    ServerConnectString\t - \"host=\"siebel://host[:port]/instance/appMgr_lang\"");
            Console.WriteLine("    LocalConnectString\t - \"lang=\"LANG_CODE\" \"cfg=\"filename.cfg, DataSource\"");
            Console.WriteLine("    filename.cfg\t - Path to Siebel configuration file");
            Console.WriteLine("   DataSource\t - section name in Siebel configuration file");
            Console.WriteLine("");
            Console.WriteLine("Example for server mode:");
            Console.WriteLine("  Siebel_DataControl \"lang='ENU' host='siebel://serv1:2321/SBA_84/FINSObjMgr_rus'");
            Console.WriteLine("Example for local mode:");
            Console.WriteLine("  Siebel_DataControl \"lang='ENU' cfg='C:\\Siebel\\15.0.0.0.0\\Client\\BIN\\enu\fins.cfg,ServerDataSrc'");
            Console.WriteLine("");
            ConsoleColor cl = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("https://docs.oracle.com/cd/B40099_02/books/OIRef/OIRefProgramming45.html#wp1006173");
            Console.ForegroundColor = cl;
        }
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                printUsage();
                Environment.Exit(-1);
            }

            string connString = String.Join(" ", args);
            connString = connString.Replace('\'', '"');

            SiebelDataControl app = new SiebelDataControl();
            app.EnableExceptions(0);

            //ConnectMode.Server
            //connString = "host=\"siebel://serv1:2321/SBA_84/FINSObjMgr_rus\"";
            connString = "lang=\"ENU\" host=\"siebel://serv1:2321/SBA_84/FINSObjMgr_enu\"";

            //ConnectMode.Local
            //connString = "lang=\"ENU\" cfg=\"C:\\Siebel\\15.0.0.0.0\\Client\\BIN\\enu\\fins.cfg, ServerDataSrc\"";            

            Console.WriteLine("Try establish connection to Siebel with connectString: "+connString);
            bool success = app.Login(connString, "SADMIN", "SADMIN");
            if (!success)
            {
                int ErrCode = app.GetLastErrCode(); //4122 error
                string ErrMsg = app.GetLastErrText();
                string err = String.Format("Connection failed. ErrCode: {0}, ErrMessage: {1}", ErrCode, ErrMsg);
                Trace.WriteLine(err);
                Console.WriteLine(err);
                Environment.Exit(ErrCode);
            }
            Console.WriteLine("Successfully connected.");
        }
    }
}
