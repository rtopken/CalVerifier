using System;

namespace CalVerifier
{
    public class Globals
    {
        public static bool bListMode = false;
        public static bool bMoveItems = false;
        public static bool bVerbose = false;
        public static string strListFile = "";
        public static string strTxtOutFile = "";
        public static string strCSVOutFile = "";
        public static string[] rgstrMBX;


        public static void ResetGlobals()
        {
            bListMode = false;
            bMoveItems = false;
            bVerbose = false;
            strListFile = "";
            strTxtOutFile = "";
            strCSVOutFile = "";
        }
    }
}