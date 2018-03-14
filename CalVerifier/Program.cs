using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System.IO;
using static CalVerifier.Globals;
using static CalVerifier.Utils;
using static CalVerifier.PropSet;
using static CalVerifier.Process;

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
                            strListFile = args[i + 1];
                            if (File.Exists(strListFile))
                            {
                                bListMode = true;
                            }
                            else
                            {
                                Console.WriteLine("Could not find the file " + strListFile + ".");
                                ShowHelp();
                                return;
                            }
                        }
                    }

                    if (args[i].ToUpper() == "-M" || args[i].ToUpper() == "/M") // move mode to move problem items out to the CalVerifier folder
                    {
                        bMoveItems = true;
                    }

                    if (args[i].ToUpper() == "-V" || args[i].ToUpper() == "/V") // include tracing, verbose mode.
                    {
                        bVerbose = true;
                    }

                    if (args[i].ToUpper() == "-?" || args[i].ToUpper() == "/?") // display command switch help
                    {
                        ShowHelp();
                        return;
                    }
                }
            }

            exService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            exService.UseDefaultCredentials = false;

            if (bVerbose)
            {
                exService.TraceEnabled = true;
                exService.TraceFlags = TraceFlags.All;
            }

            ShowInfo();

            if (bListMode)
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

            exService.Credentials = new WebCredentials(strAcct, strPwd);
            exService.AutodiscoverUrl(strAcct, RedirectionUrlValidationCallback);

            if (bListMode) // List mode
            {
                rgstrMBX = File.ReadAllLines(strListFile);
                foreach (string strSMTP in rgstrMBX)
                {
                    CreateLogFile();
                    LogInfo();
                    NameResolutionCollection ncCol = exService.ResolveName(strSMTP, ResolveNameSearchLocation.DirectoryOnly, true);
                    strDisplayName = ncCol[0].Contact.DisplayName;
                    DisplayAndLog("Processing Calendar for " + strDisplayName);

                    exService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, strSMTP);
                    CalItems = GetCalItems(exService);
                    strSMTPAddr = strSMTP.ToUpper();
                    if (CalItems != null)
                    {
                        string strCount = CalItems.TotalCount.ToString();
                        DisplayAndLog("Found " + strCount + " items");
                        DisplayAndLog("");
                        Console.Write("Processing items...");
                    }
                    else
                    {
                        return; // could not connect, error is displayed to user already.
                    }

                    int i = 0;
                    foreach (Appointment appt in CalItems)
                    {
                        i++;
                        if (i % 50 == 0)
                            Console.Write(".");
                        ProcessItem(appt);
                    }
                    DisplayAndLog("===============================================================");
                    DisplayAndLog("Found " + iErrors.ToString() + " errors and " + iWarn.ToString() + " warnings.");
                    DisplayAndLog("===============================================================");
                    outLog.Close();
                    if (File.Exists(strAppPath + strSMTPAddr + "_CalVerifier.log"))
                    {
                        File.Delete(strAppPath + strSMTPAddr + "_CalVerifier.log");
                    }
                    File.Move(strLogFile, strAppPath + strSMTPAddr + "_CalVerifier.log");
                }
            }
            else // single mailbox mode
            {
                CreateLogFile();
                LogInfo();
                NameResolutionCollection ncCol = exService.ResolveName(strAcct, ResolveNameSearchLocation.DirectoryOnly, true);
                if (ncCol[0].Contact != null)
                {
                    strDisplayName = ncCol[0].Contact.DisplayName;
                    DisplayAndLog("Processing Calendar for " + strDisplayName);
                }
                else
                {
                    DisplayAndLog("Processing Calendar for " + strAcct);
                }

                strSMTPAddr = ncCol[0].Mailbox.Address.ToUpper();

                CalItems = GetCalItems(exService);
                if (CalItems != null)
                {
                    string strCount = CalItems.TotalCount.ToString();
                    DisplayAndLog("Found " + strCount + " items");
                    DisplayAndLog("");
                    Console.Write("Processing items...");
                }
                else
                {
                    return;  // could not connect, error is displayed to user already.
                }

                int i = 0;
                foreach (Appointment appt in CalItems)
                {
                    i++;
                    if (i % 50 == 0)
                        Console.Write(".");
                    ProcessItem(appt);
                }
                DisplayAndLog("===============================================================");
                DisplayAndLog("Found " + iErrors.ToString() + " errors and " + iWarn.ToString() + " warnings.");
                DisplayAndLog("===============================================================");
                outLog.Close();
                if (File.Exists(strAppPath + strSMTPAddr + "_CalVerifier.log"))
                {
                    File.Delete(strAppPath + strSMTPAddr + "_CalVerifier.log");
                }
                File.Move(strLogFile, strAppPath + strSMTPAddr + "_CalVerifier.log");
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
            DoProps(ref propSet);
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
