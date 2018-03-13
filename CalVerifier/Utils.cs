using System;
using System.IO;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace CalVerifier
{
    public class Utils
    {
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

        public static ExchangeService exService;

        public static ExchangeService GetExService()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            return service;
        }

        public static void LogLine(string strLine)
        {
            Globals.outLog.WriteLine(strLine);
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