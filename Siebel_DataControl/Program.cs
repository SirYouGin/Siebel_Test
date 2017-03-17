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
        static void Main(string[] args)
        {
            SiebelDataControl app = new SiebelDataControl();
            app.EnableExceptions(0);
            string connString = "host=\"siebel://serv1:2321/SBA_84/FINSObjMgr_rus\"";
            Console.WriteLine("Try establish connection to Siebel with connectString: "+connString);
            bool success = app.Login(connString, "SADMIN", "SADMIN");
            if (!success)
            {
                int ErrCode = app.GetLastErrCode();
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
