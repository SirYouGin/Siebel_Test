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
        enum SiebelViewModeConstants { SalesRepView, ManagerView, PersonalView, AllView, NoneSetView, OrgView, ContactView, SubOrgView, GroupView, CatalogView };
        enum SiebelQueryConstants {ForwardBackward, ForwardOnly};

        private static SiebelDataControl app;

        private static int ErrorCode = 0;
        private static void checkError()
        {
            ErrorCode = app.GetLastErrCode();
            if (ErrorCode != 0)
            {
                string s = "ErrCode: " + ErrorCode + " ErrMsg: " + app.GetLastErrText();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(s);
                Console.ResetColor();
                Trace.WriteLine(s);                
                Environment.Exit(ErrorCode);
            }
        }
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
            Console.WriteLine("  Siebel_DataControl lang='ENU' host='siebel://serv1:2321/SBA_84/FINSObjMgr_rus'");
            Console.WriteLine("Example for local mode:");
            Console.WriteLine("  Siebel_DataControl lang='ENU' cfg='C:\\Siebel\\15.0.0.0.0\\Client\\BIN\\enu\fins.cfg,ServerDataSrc'");
            Console.WriteLine("");
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("https://docs.oracle.com/cd/B40099_02/books/OIRef/OIRefProgramming45.html#wp1006173");
            Console.ResetColor();
        }
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                printUsage();
                Environment.Exit(-1);
            }

            Dictionary<string, string> fields = new Dictionary<string, string>();

            string connString = String.Join(" ", args);
            connString = connString.Replace('\'', '"');

            app = new SiebelDataControl();
            app.EnableExceptions(0);

            // Examples:

            //ConnectMode.Server
            //connString = "host=\"siebel://serv1:2321/SBA_84/FINSObjMgr_rus\"";
            //connString = "host=\"siebel://serv1:2321/SBA_84/FINSeSalesObjMgr_rus\"";

            //ConnectMode.Local
            //connString = "lang=\"ENU\" cfg=\"C:\\Siebel\\15.0.0.0.0\\Client\\BIN\\enu\\fins.cfg, ServerDataSrc\"";            

            Console.Write("try connect to Siebel with connectString: \"{0}\"...",connString);
            bool success = app.Login(connString, "SADMIN", "SADMIN");              
            Console.WriteLine("{0}\n", (success ? "OK" :"Failed"));

            checkError();

            SiebelBusObject bo = app.GetBusObject("Contact"); checkError();

            SiebelBusComp bc = bo.GetBusComp("Contact"); checkError();

            bc.ClearToQuery(); checkError();

            bc.SetViewMode(Convert.ToInt16(SiebelViewModeConstants.AllView)); checkError();

            bc.SetSearchSpec("First Name", "*"); checkError();

            bc.ExecuteQuery(Convert.ToInt16(SiebelQueryConstants.ForwardOnly)); checkError();

            bool isRecord = bc.FirstRecord(); checkError();

            string fname;                    

            while (isRecord)
            {
                fname = "Id"; fields.Add(fname, bc.GetFieldValue(fname)); checkError();
                fname = "First Name"; fields.Add(fname, bc.GetFieldValue(fname)); checkError();
                fname = "Last Name"; fields.Add(fname, bc.GetFieldValue(fname)); checkError();
                Console.WriteLine("Id=" + fields["Id"] + " FirstName=" + fields["First Name"] + " Last Name=" + fields["Last Name"]);
                fields.Clear();
                isRecord = bc.NextRecord(); checkError();
            }

            Console.Write("\nDisconnect from Siebel...");
            success = app.Logoff();
            Console.WriteLine("{0}\n"+ (success ? "OK":"Failed"));

            Console.WriteLine("Press any key to close application.");
            Console.ReadKey();
        }
    }
}
