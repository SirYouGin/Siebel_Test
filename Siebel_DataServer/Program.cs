using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using SiebelDataServer;

namespace Siebel_DataServer
{
    class Program
    {
        private static SiebelDataServer.SiebelApplication app;

        private static short ErrorCode = -1;
        private static void checkError()
        {
            if (ErrorCode != 0)
            {
                string s = "ErrCode: "+ ErrorCode+" ErrMsg: "+app.GetLastErrText();
                if (Enum.IsDefined(typeof(SiebelEnumErrCode), (Int32)ErrorCode))
                {
                    s = s + " Siebel Error: " + (SiebelEnumErrCode)ErrorCode;
                }
                Console.WriteLine(s);
                Trace.WriteLine(s);
                
                Environment.Exit(ErrorCode);                                              
            }
        }


        private static void printUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("\tSiebel_DataServer \"filename.cfg, DataSource\"");
            Console.WriteLine("\t\tfilename.cfg\t - Path to Siebel configuration file");
            Console.WriteLine("\t\tDataSource\t - section name in Siebel configuration file");            
            Console.WriteLine("\t\t\t\t ");
            Console.WriteLine("Example:");
            Console.WriteLine("\tSiebel_DataServer \"C:\\Siebel\\15.0.0.0.0\\Client\\BIN\\enu\\fins.cfg, ServerDataSrc\"");
        }


        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                printUsage();
                Environment.Exit(-1);
            }

            Dictionary<string, string> fields = new Dictionary<string, string>(); 
           
            Type SiebelAppType = Type.GetTypeFromProgID("SiebelDataServer.ApplicationObject",true);
            app = (SiebelApplication)Activator.CreateInstance(SiebelAppType);

            string cfgpath = args[0]; //@"C:\Siebel\15.0.0.0.0\Client\BIN\enu\fins.cfg, ServerDataSrc";

            Console.WriteLine("read Siebel configuration file \"{0}\"...",cfgpath);

            app.LoadObjects(cfgpath, ref ErrorCode); checkError();

            Console.WriteLine("try connect to Siebel...");

            app.Login("SADMIN", "SADMIN", ref ErrorCode); checkError();

            Console.WriteLine("Successfully connected.\n");

            SiebelBusObject bo = app.GetBusObject("Contact", ref ErrorCode); checkError();

            SiebelBusComp bc = bo.GetBusComp("Contact", ref ErrorCode); checkError();

            bc.ClearToQuery(ref ErrorCode); checkError();

            bc.SetViewMode(Convert.ToInt16(SiebelViewModeConstants.AllView), ref ErrorCode); checkError();

            bc.SetSearchSpec("First Name", "*", ref ErrorCode); checkError();

            bc.ExecuteQuery(Convert.ToBoolean(SiebelQueryConstants.ForwardOnly), ref ErrorCode); checkError();

            bool isRecord = bc.FirstRecord(ref ErrorCode); checkError();

            string fname;

            while (isRecord)
            {
                fname = "Id";  fields.Add(fname, bc.GetFieldValue(fname, ref ErrorCode)); checkError();
                fname = "First Name"; fields.Add(fname, bc.GetFieldValue(fname, ref ErrorCode)); checkError();
                fname = "Last Name"; fields.Add(fname, bc.GetFieldValue(fname, ref ErrorCode)); checkError();
                Console.WriteLine("Id=" + fields["Id"] + " FirstName=" + fields["First Name"] + " Last Name=" + fields["Last Name"]);
                fields.Clear();
                isRecord = bc.NextRecord(ref ErrorCode); checkError();
            }

            Console.WriteLine("\nDisconnect from Siebel.");
        }

    }
}
