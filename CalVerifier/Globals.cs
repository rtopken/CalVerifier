using System;
using System.IO;

namespace CalVerifier
{
    public class Globals
    {
        public static bool bListMode = false;
        public static bool bMoveItems = false;
        public static bool bVerbose = false;
        public static string strListFile = "";
        public static string[] rgstrMBX;
        public static string strAppPath = AppDomain.CurrentDomain.BaseDirectory;
        public static int iErrors = 0;
        public static DateTime dtMin = DateTime.Parse("01/01/1601 00:00");
        public static DateTime dtMax = DateTime.Parse("12/31/4500 11:59"); 
        public static DateTime dtNone = DateTime.Parse("01/01/4501 00:00");
        public static string[] rgstrProxyAddresses;
        public static string strDisplayName = "";
        public static string strSMTPAddr = "";

        public static StreamWriter outLog = new StreamWriter(strAppPath + "CalVerifier.log");


        public static void ResetGlobals()
        {
            bListMode = false;
            bMoveItems = false;
            bVerbose = false;
            strListFile = "";
            strAppPath = "";
            rgstrMBX = null;
        }
    }
}