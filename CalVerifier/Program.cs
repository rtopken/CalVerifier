using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.IO;

namespace CalVerifier
{
    class Program
    {
        static void Main(string[] args)
        {
            string strAcct = "";
            string strPwd = "";
            FindItemsResults<Item> CalItems = null;

            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].ToUpper() == "-L" || args[i].ToUpper() == "/L") // list mode pulls in a list of SMTP addresses - will connect to each of them.
                    {
                        if (args[i + 1].Length > 0)
                        {
                            Globals.strListFile = args[i + 1];
                            if (File.Exists(Globals.strListFile))
                            {
                                Globals.bListMode = true;
                            }
                            else
                            {
                                Console.WriteLine("Could not find the file " + Globals.strListFile + ".");
                                Utils.ShowHelp();
                                goto Exit;
                            }
                        }
                    }

                    if (args[i].ToUpper() == "-M" || args[i].ToUpper() == "/M") // move mode to move problem items out to the CalVerifier folder
                    {
                        Globals.bMoveItems = true;
                    }

                    if (args[i].ToUpper() == "-V" || args[i].ToUpper() == "/V") // include tracing, verbose mode.
                    {
                        Globals.bVerbose = true;
                    }

                    if (args[i].ToUpper() == "-?" || args[i].ToUpper() == "/?") // display command switch help
                    {
                        Utils.ShowHelp();
                        goto Exit;
                    }
                }
            }

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.UseDefaultCredentials = false;

            if (Globals.bVerbose)
            {
                service.TraceEnabled = true;
                service.TraceFlags = TraceFlags.All;
            }

            Utils.ShowInfo();

            if (Globals.bListMode)
            {
                Console.Write("Enter the SMTP address of the ServiceAccount: ");
            }
            else
            {
                Console.Write("Enter the SMTP address of the Mailbox: ");
            }

            strAcct = Console.ReadLine();
            Console.Write("Enter the password for {0}: ", strAcct);

            // use below while loop to mask the password while reading it in
            bool bEnter = true;
            int iPwdChars = 0;
            while (bEnter)
            {
                ConsoleKeyInfo ckiKey = Console.ReadKey(true);
                if (ckiKey.Key == ConsoleKey.Enter)
                {
                    bEnter = false;
                }
                else if (ckiKey.Key == ConsoleKey.Backspace)
                {
                    if (strPwd.Length >= 1)
                    {
                        int oldLength = strPwd.Length;
                        strPwd = strPwd.Substring(0, oldLength - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    strPwd = strPwd + ckiKey.KeyChar.ToString();
                    iPwdChars++;
                    Console.Write('*');
                }
            }

            Console.WriteLine();

            service.Credentials = new WebCredentials(strAcct, strPwd);
            service.AutodiscoverUrl(strAcct, RedirectionUrlValidationCallback);

            if (Globals.bListMode) // List mode
            {
                Globals.rgstrMBX = File.ReadAllLines(Globals.strListFile);
                foreach (string strSMTP in Globals.rgstrMBX)
                {
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, strSMTP);
                    CalItems = GetCalItems();
                    Console.WriteLine(strSMTP);
                    Console.WriteLine("Found {0} items", CalItems.TotalCount);
                    Console.WriteLine("");

                    foreach (Appointment appt in CalItems)
                    {
                        Process.ProcessItem(appt);
                    }
                }
            }
            else // single mailbox mode
            {
                CalItems = GetCalItems();
                Console.WriteLine(strAcct);
                Console.WriteLine("Found {0} items", CalItems.TotalCount);
                Console.WriteLine("");

                foreach (Appointment appt in CalItems)
                {
                    Process.ProcessItem(appt);
                }
            }

            FindItemsResults<Item> GetCalItems()
            {
                // Here's where it will do the connect to the user / Calendar
                Folder fldCal = Folder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());

                // if we're in then we get here
                // creating a view with props to request / collect
                ItemView cView = new ItemView(int.MaxValue);
                List<ExtendedPropertyDefinition> propSet = new List<ExtendedPropertyDefinition>();
                PropSet.DoProps(ref propSet);
                cView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
                foreach (PropertyDefinitionBase pdbProp in propSet)
                {
                    cView.PropertySet.Add(pdbProp);
                }

                // now go get the items
                FindItemsResults<Item> cAppts = fldCal.FindItems(cView);
                return cAppts;
            }

            Exit:
            // Exit the app...
            //Console.Write("\r\nExiting the program.");
            //Console.ReadLine();
            return;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
