﻿using System;
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
        public enum SiebelEnumErrCode
        {
            SSGenErrInternal = 4096,
            SSELErrAppletNotActive = 4097,
            SSELErrAppletHasNoBusComp = 4098,
            SSELErrBadNewWhere = 4099,
            SSELErrBadViewModeNum = 4100,
            SSELErrControlNotInApplet = 4101,
            SSELErrGotoViewActiveBusObj = 4102,
            SSELErrInactiveApplet = 4103,
            SSELErrMethodNotSupported = 4104,
            SSELErrNoAppletName = 4105,
            SSELErrNoControlName = 4106,
            SSELErrNoVariableName = 4107,
            SSELErrNoViewName = 4108,
            SSELErrNotAssocList = 4109,
            SSELErrOutOfScope = 4110,
            SSELErrUnknownApplet = 4111,
            SSELErrUnknownControl = 4112,
            SSELErrUnknownView = 4113,
            SSOleErrOleDisabled = 4114,
            SSOleErrInternal = 4115,
            SSOleErrMemory = 4116,
            SSOleErrMissingAppletArg = 4117,
            SSOleErrMissingControlArg = 4118,
            SSOleErrMissingStringArg = 4119,
            SSOleErrMissingOrInvalidConfigFile = 4120,
            SSOleErrStringArrayExpected = 4121,
            SSOleErrUnableToLogin = 4122,
            SSOMErrActRowChanged = 4123,
            SSOMErrActiveRow = 4124,
            SSOMErrAddAssoc = 4125,
            SSOMErrAssocCurRow = 4126,
            SSOMErrAssocList = 4127,
            SSOMErrBOBusComp = 4128,
            SSOMErrBOQuerySyntax = 4129,
            SSOMErrBoundedPick = 4130,
            SSOMErrBusObjProp = 4131,
            SSOMErrConstrBusComp = 4132,
            SSOMErrContextSyntax = 4133,
            SSOMErrDeleteRecord = 4134,
            SSOMErrEOWS = 4135,
            SSOMErrEmptyMVGParent = 4136,
            SSOMErrEnd = 4137,
            SSOMErrExec = 4138,
            SSOMErrGetBusComp = 4139,
            SSOMErrGetContext = 4140,
            SSOMErrGetMVGroup = 4141,
            SSOMErrGetQuerySpec = 4142,
            SSOMErrGotoBookmark = 4143,
            SSOMErrHome = 4144,
            SSOMErrInvalidLogin = 4145,
            SSOMErrMVGCopyRec = 4146,
            SSOMErrMVGQuery = 4147,
            SSOMErrMVSrcInactive = 4148,
            SSOMErrNewRecIdValue = 4149,
            SSOMErrNewRecord = 4150,
            SSOMErrNextRecord = 4151,
            SSOMErrNextSet = 4152,
            SSOMErrNoAssocList = 4153,
            SSOMErrNoAssocMVGAllowed = 4154,
            SSOMErrNoDelete = 4155,
            SSOMErrNoInsert = 4156,
            SSOMErrParentBCRequired = 4157,
            SSOMErrParentMVGEmpty = 4158,
            SSOMErrPickCurRow = 4159,
            SSOMErrPopupRequired = 4160,
            SSOMErrPriorRecord = 4161,
            SSOMErrPriorSet = 4162,
            SSOMErrScrollWorkSet = 4163,
            SSOMErrSearchNotSupp = 4164,
            SSOMErrSetActiveRow = 4165,
            SSOMErrSetContext = 4166,
            SSOMErrSetFieldValue = 4167,
            SSOMErrSetMVGPrimaryId = 4168,
            SSOMErrSortNotSupp = 4169,
            SSOMErrUpdRecord = 4170,
            SSOMErrUserAbort = 4171,
            SSSqlErrAssocObj = 4172,
            SSSqlErrBOF = 4173,
            SSSqlErrDupConflict = 4174,
            SSSqlErrDupConflict2 = 4175,
            SSSqlErrEOF = 4176,
            SSSqlErrEndNoSort = 4177,
            SSSqlErrEndTrx = 4178,
            SSSqlErrEvalFieldExpr = 4179,
            SSSqlErrFieldExist = 4180,
            SSSqlErrFieldNoQBE = 4181,
            SSSqlErrFieldReadOnly = 4182,
            SSSqlErrFieldSearch = 4183,
            SSSqlErrLogon = 4184,
            SSSqlErrRecordDeleted = 4185,
            SSSqlErrReqField = 4186,
            SSSqlErrRsltsDiscarded = 4187,
            SSELErrBadAssocWhere = 4188,
            SSELErrBadDrilldownName = 4189,
            SSCorbaErrCorbaDisabled = 4190,
            SSCorbaErrInternal = 4191,
            SSCorbaErrMemory = 4192,
            SSCorbaErrMissingAppletArg = 4193,
            SSCorbaErrMissingControlArg = 4194,
            SSCorbaErrMissingStringArg = 4195,
            SSCorbaErrMissingOrInvalidConfigFile = 4196,
            SSCorbaErrStringArrayExpected = 4197,
            SSCorbaErrUnableToLogin = 4198,
            SSELErrBusCompNameRequired = 4199,
            SSELErrCannotCreateService = 4200,
            SSELErrMethodNameRequired = 4201,
            SSELErrServiceNameRequired = 4202,
            SSELErrUserDefinedError = 4203
        };

        private static SiebelDataControl app;

        private static int ErrorCode = 0;
        private static void checkError()
        {
            ErrorCode = app.GetLastErrCode();
            if (ErrorCode != 0)
            {
                string s = "ErrCode: " + ErrorCode + " ErrMsg: " + app.GetLastErrText();
                if (Enum.IsDefined(typeof(SiebelEnumErrCode), (Int32)ErrorCode))
                {
                    s = s + " Siebel Desc: " + (SiebelEnumErrCode)ErrorCode;
                }
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(s);
                Console.ResetColor();
                Trace.WriteLine(s);

                Console.WriteLine("\nPress any key to close application.");
                Console.ReadKey();

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
            connString = "host=\"siebel://serv1:2321/SBA_84/FINSObjMgr_rus\"";
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
            Console.WriteLine("{0}\n", (success ? "OK":"Failed"));

            Console.WriteLine("Press any key to close application.");
            Console.ReadKey();
        }
    }
}