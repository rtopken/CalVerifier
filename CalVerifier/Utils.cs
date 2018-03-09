using System;
using System.IO;
using System.Text;

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
    }
}