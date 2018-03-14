using System;
using System.IO;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Diagnostics;
using static CalVerifier.Globals;

namespace CalVerifier
{
    public class Utils
    {
        // So I can get to the service object from wherever...
        public static ExchangeService exService;

        public static string strMrMAPIFile;

        public static void ShowInfo()
        {
            Console.WriteLine("");
            Console.WriteLine("===========");
            Console.WriteLine("CalVerifier");
            Console.WriteLine("===========");
            Console.WriteLine("Checks Calendars for potential problem items and reports them.\r\n");
        }
        public static void ShowHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("CalVerifier [-L <filename>] [-M] [-V] [-?]");
            Console.WriteLine("");
            Console.WriteLine("-L   [List mode. Requires a file with SMTP addresses of mailboxes to check - one SMTP address per line.]");
            Console.WriteLine("-M   [Move mode. Will move problem items out to a folder called CalVerifier.]");
            Console.WriteLine("-V   [Verbose. Will output tracing information.]");
            Console.WriteLine("-?   [Shows this usage information.]");
            Console.WriteLine("");
        }

        public static void LogInfo()
        {
            LogLine("===========");
            LogLine("CalVerifier");
            LogLine("===========");
            LogLine("Checks Calendars for potential problem items and reports them.\r\n");
        }

        public static void LogLine(string strLine)
        {
            outLog.WriteLine(strLine);
        }

        public static void DisplayAndLog(string strLine)
        {
            Console.WriteLine(strLine);
            outLog.WriteLine(strLine);
        }

        // create a hex blob text file for use with MrMAPI
        public static void CreateHexFile(string strHex, string strName)
        {
            strMrMAPIFile = strAppPath + strName;
            if (File.Exists(strMrMAPIFile))
            {
                File.Delete(strMrMAPIFile);
            }
            StreamWriter swFile = new StreamWriter(strMrMAPIFile);
            swFile.WriteLine(strHex);
            swFile.Close();
        }

        // run the MrMAPI app to get item props
        public static void RunMrMAPI(string strSwitches)
        {
            // Setup and launch MrMAPI
            string strAppPath = GetMrMAPIPath();
            if (File.Exists(strAppPath))
            {
                System.Diagnostics.Process mrMAPI = new System.Diagnostics.Process();
                mrMAPI.StartInfo.FileName = (strAppPath);
                mrMAPI.StartInfo.Arguments = (strSwitches);
                mrMAPI.StartInfo.UseShellExecute = true;
                mrMAPI.StartInfo.CreateNoWindow = true;
                mrMAPI.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                mrMAPI.Start();
                mrMAPI.WaitForExit();
            }
            return;
        }

        // Get the path to MrMAPI
        public static string GetMrMAPIPath()
        {
            object oRegVal;
            string strRegVal;
            string strSize = IntPtr.Size.ToString();
            string mapiBitness = "x86";

            // Check for ClickToRun config first...
            if (Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun", "Platform", null) != null)
            {
                oRegVal = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun", "Platform", null);
                strRegVal = oRegVal.ToString();
                if (strRegVal.Contains("x64"))
                {
                    mapiBitness = "x64";
                }
            }
            // check Office 15.0 location
            else if (Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Outlook", "Bitness", null) != null)
            {
                oRegVal = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Outlook", "Bitness", null);
                strRegVal = oRegVal.ToString();
                if (strRegVal.Contains("x64"))
                {
                    mapiBitness = "x64";
                }
            }
            // check Office 16.0 location
            else if (Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Outlook", "Bitness", null) != null)
            {
                oRegVal = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Outlook", "Bitness", null);
                strRegVal = oRegVal.ToString();
                if (strRegVal.Contains("x64"))
                {
                    mapiBitness = "x64";
                }
            }
            // I "think" Outlook 2013 would populate this with MSI builds - so another way to check if others above fail
            else if (IntPtr.Size == 8 && Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Messaging Subsystem\MSMapiApps", "outlook.exe", null) != null)
            {
                mapiBitness = "x64";
            }

            return System.IO.Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), mapiBitness, "MrMAPI.exe");
        }

        // Check date/time values against boundary values.
        // return TRUE if the time is no good.
        public static bool TimeCheck(DateTime dtCheck)
        {
            int iComp = 0;
            // less than 0 >> t1 earlier than t2
            // zero >> t1 same as t2
            //greater than 0 >> t1 is later than t2

            iComp = DateTime.Compare(dtCheck, Globals.dtMin);
            if (iComp <= 0)
            {
                return true;
            }

            iComp = DateTime.Compare(dtCheck, Globals.dtMax);
            if (iComp >= 0)
            {
                return true;
            }

            iComp = DateTime.Compare(dtCheck, Globals.dtNone);
            if (iComp >= 0)
            {
                return true;
            }

            if (dtCheck <= DateTime.MinValue || dtCheck >= DateTime.MaxValue)  // these are probably different from the Outlook boundary values
            {
                return true;
            }

            return false;
        }

    }
}