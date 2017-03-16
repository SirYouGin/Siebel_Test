using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
//using SiebelBusObjectInterfaces;
using SiebelDataServer;

namespace Siebel_Test
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

        [STAThread]
        static void Main(string[] args)
        {

            Dictionary<string, string> fields = new Dictionary<string, string>(); 

            //Type SiebelAppType = Type.GetTypeFromProgID("SiebelDataControl.SiebelDataControl");
            //SiebelDataControl SiebelApp = new SiebelDataControl();
            //bool ret = SiebelApp.Login("host=\"siebel://localhost:2321/SBA_84/FINSObjMgr_enu\"", "SADMIN", "SADMIN");            

            Type SiebelAppType = Type.GetTypeFromProgID("SiebelDataServer.ApplicationObject",true);
            app = (SiebelApplication)Activator.CreateInstance(SiebelAppType);

            app.LoadObjects(@"C:\Siebel\15.0.0.0.0\Client\BIN\enu\fins.cfg, ServerDataSrc", ref ErrorCode); checkError();

            app.Login("SADMIN", "SADMIN", ref ErrorCode); checkError();

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
                Trace.WriteLine("Id=" + fields["Id"] + " FirstName=" + fields["First Name"] + " Last Name=" + fields["Last Name"]);
                fields.Clear();
                isRecord = bc.NextRecord(ref ErrorCode); checkError();
            }           

        }
    }
}
