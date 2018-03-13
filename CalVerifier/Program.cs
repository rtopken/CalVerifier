using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System.IO;

namespace CalVerifier
{
    class Program
    {
        static void Main(string[] args)
        {
            string strAcct = "";
            string strPwd = "";
            
            List<string> strProxyAddresses = new List<string>();
            
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
                                return;
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
                        return;
                    }
                }
            }

            Utils.exService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            Utils.exService.UseDefaultCredentials = false;

            if (Globals.bVerbose)
            {
                Utils.exService.TraceEnabled = true;
                Utils.exService.TraceFlags = TraceFlags.All;
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

            Utils.exService.Credentials = new WebCredentials(strAcct, strPwd);
            Utils.exService.AutodiscoverUrl(strAcct, RedirectionUrlValidationCallback);

            

            if (Globals.bListMode) // List mode
            {
                Globals.rgstrMBX = File.ReadAllLines(Globals.strListFile);
                foreach (string strSMTP in Globals.rgstrMBX)
                {
                    NameResolutionCollection ncCol = Utils.exService.ResolveName(strSMTP, ResolveNameSearchLocation.DirectoryOnly, true);
                    Globals.strDisplayName = ncCol[0].Contact.DisplayName;
                    Console.WriteLine("Processing Calendar for " + Globals.strDisplayName);

                    Utils.exService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, strSMTP);
                    CalItems = GetCalItems(Utils.exService);
                    Globals.strSMTPAddr = strSMTP.ToUpper();
                    if (CalItems != null)
                    {
                        Console.WriteLine("Found {0} items", CalItems.TotalCount);
                        Console.WriteLine("");
                    }
                    else
                    {
                        return; // could not connect, error is displayed to user already.
                    }

                    foreach (Appointment appt in CalItems)
                    {
                        Process.ProcessItem(appt);
                    }
                }
            }
            else // single mailbox mode
            {
                NameResolutionCollection ncCol = Utils.exService.ResolveName(strAcct, ResolveNameSearchLocation.DirectoryOnly, true);
                Globals.strDisplayName = ncCol[0].Contact.DisplayName;
                Console.WriteLine("Processing Calendar for " + Globals.strDisplayName);

                Globals.strSMTPAddr = ncCol[0].Mailbox.Address.ToUpper();

                CalItems = GetCalItems(Utils.exService);
                if (CalItems != null)
                {
                    Console.WriteLine("Found {0} items", CalItems.TotalCount);
                    Console.WriteLine("");
                }
                else
                {
                    return;  // could not connect, error is displayed to user already.
                }

                foreach (Appointment appt in CalItems)
                {
                    Process.ProcessItem(appt);
                }
            }
        }

        public static FindItemsResults<Item> GetCalItems(ExchangeService service)
        {
            Folder fldCal = null;
            try
            {
                // Here's where it will do the connect to the user / Calendar
                fldCal = Folder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
            }
            catch
            {
                Console.WriteLine("Could not connect to this user's mailbox or calendar.");
                return null;
            }

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
