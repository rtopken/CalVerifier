using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;
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
            string strTenant = "";
            string strEmailAddr = "";

            List<Item> CalItems = null;

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
                Console.Write("Press <ENTER> to enter credentials for the ServiceAccount.");
            }
            else
            {
                Console.Write("Press <ENTER> to enter credentials for the Mailbox.");
            }

            Console.ReadLine();
            Console.WriteLine();
            
            AuthenticationResult authResult = GetToken();
            exService.Credentials = new OAuthCredentials(authResult.AccessToken);
            strAcct = authResult.UserInfo.DisplayableId;
            strTenant = strAcct.Split('@')[1];
            exService.Url = new Uri(strSrvURI + "/ews/exchange.asmx");

            NameResolutionCollection ncCol = null;

            // Should only do this if the switch was set.
            if (bMoveItems)
            {
                CreateErrFld();
            }

            if (bListMode) // List mode
            {
                rgstrMBX = File.ReadAllLines(strListFile);
                foreach (string strSMTP in rgstrMBX)
                {
                    CreateLogFile();
                    LogInfo();
                    ncCol = DoResolveName(strSMTP);
                    if (ncCol == null)
                    {
                        // Didn't get a NameResCollection, so error out.
                        Console.WriteLine("");
                        Console.WriteLine("Exiting the program.");
                        return;
                    }

                    if (ncCol[0].Contact != null)
                    {
                        strDisplayName = ncCol[0].Contact.DisplayName;
                        DisplayAndLog("Processing Calendar for " + strDisplayName);
                    }
                    else
                    {
                        DisplayAndLog("Processing Calendar for " + strSMTP);
                    }

                    exService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, strSMTP);
                    CalItems = GetCalItems(exService);
                    strSMTPAddr = strSMTP.ToUpper();
                    if (CalItems != null)
                    {
                        string strCount = CalItems.Count.ToString();
                        DisplayAndLog("Found " + strCount + " items");
                        DisplayAndLog("");
                        Console.Write("Processing items...");
                    }
                    else
                    {
                        return; // could not connect, error is displayed to user already.
                    }

                    int i = 0;
                    int n = 0;
                    foreach (Appointment appt in CalItems)
                    {
                        i++;
                        if (i % 5 == 0)
                        {
                            Console.SetCursorPosition(0, Console.CursorTop);
                            Console.Write("");
                            Console.Write(cSpin[n % 4]);
                            n++;
                        }
                        ProcessItem(appt);
                        iCheckedItems++;
                    }
                    DisplayAndLog("\r\n");
                    DisplayAndLog("===============================================================");
                    DisplayAndLog("Checked " + iCheckedItems.ToString() + " items.");
                    DisplayAndLog("Found " + iErrors.ToString() + " errors and " + iWarn.ToString() + " warnings.");
                    DisplayAndLog("===============================================================");

                    outLog.Close();

                    if (File.Exists(strAppPath + strSMTPAddr + "_CalVerifier.log"))
                    {
                        File.Delete(strAppPath + strSMTPAddr + "_CalVerifier.log");
                    }
                    File.Move(strLogFile, strAppPath + strSMTPAddr + "_CalVerifier.log");

                }
                Console.WriteLine("");
                Console.WriteLine("Please check " + strAppPath + " for the CalVerifier logs.");
            }
            else // single mailbox mode
            {
                CreateLogFile();
                LogInfo();
                ncCol = DoResolveName(strAcct);
                if (ncCol == null)
                {
                    // Didn't get a NameResCollection, so error out.
                    Console.WriteLine("");
                    Console.WriteLine("Exiting the program.");
                    return;
                }

                if (ncCol[0].Contact != null)
                {
                    strDisplayName = ncCol[0].Contact.DisplayName;
                    strEmailAddr = ncCol[0].Mailbox.Address;
                    DisplayAndLog("Processing Calendar for " + strDisplayName);
                }
                else
                {
                    DisplayAndLog("Processing Calendar for " + strAcct);
                }

                strSMTPAddr = strEmailAddr.ToUpper();

                CalItems = GetCalItems(exService);
                if (CalItems != null)
                {
                    string strCount = CalItems.Count.ToString();
                    DisplayAndLog("Found " + strCount + " items");
                    DisplayAndLog("");
                    Console.WriteLine("Processing items ");
                }
                else
                {
                    return;  // could not connect, error is displayed to user already.
                }

                int i = 0;
                int n = 0;
                foreach (Appointment appt in CalItems)
                {
                    i++;
                    if (i % 5 == 0)
                    {
                        Console.SetCursorPosition(0,Console.CursorTop);
                        Console.Write("");
                        Console.Write(cSpin[n % 4]);
                        n++;
                    }
                    ProcessItem(appt);
                    iCheckedItems++;
                }
                DisplayAndLog("\r\n");
                DisplayAndLog("===============================================================");
                DisplayAndLog("Checked " + iCheckedItems.ToString() + " items.");
                DisplayAndLog("Found " + iErrors.ToString() + " errors and " + iWarn.ToString() + " warnings.");
                DisplayAndLog("===============================================================");

                outLog.Close();

                if (File.Exists(strAppPath + strSMTPAddr + "_CalVerifier.log"))
                {
                    File.Delete(strAppPath + strSMTPAddr + "_CalVerifier.log");
                }
                File.Move(strLogFile, strAppPath + strSMTPAddr + "_CalVerifier.log");

                Console.WriteLine("");
                Console.WriteLine("Please check " + strAppPath + " for " + strSMTPAddr + "_CalVerifier.log for more information.");
            }

            DisplayPrivacyInfo();
        }

        // Go get an OAuth token to use Exchange Online 
        private static AuthenticationResult GetToken()
        {
            AuthenticationResult ar = null;
            AuthenticationContext ctx = new AuthenticationContext(strAuthCommon);

            try
            {
                ar = ctx.AcquireTokenAsync(strSrvURI, strClientID, new Uri(strRedirURI), new PlatformParameters(PromptBehavior.Always)).Result;
            }
            catch (Exception Ex)
            {
                var authEx = Ex.InnerException as AdalException;
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("An error occurred during authentication with the service:");
                Console.WriteLine(authEx.HResult.ToString("X"));
                Console.WriteLine(authEx.Message);
                Console.ResetColor();
            }
            return ar;
        }

        // Go connect to the Calendar folder and get the calendar items
        public static List<Item> GetCalItems(ExchangeService service)
        {
            Folder fldCal = null;
            int iOffset = 0;
            int iPageSize = 500;
            bool bMore = true;
            List<Item> cAppts = new List<Item>();
            FindItemsResults<Item> findResults = null;

            try
            {
                // Here's where it connects to the Calendar
                fldCal = Folder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
            }
            catch (ServiceResponseException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("");
                Console.WriteLine("Could not connect to this user's mailbox or calendar.");
                Console.WriteLine(ex.Message);
                Console.ResetColor();
                return null;
            }

            // if we're in then we get here
            // creating a view with props to request / collect
            ItemView cView = new ItemView(iPageSize, iOffset, OffsetBasePoint.Beginning);
            List<ExtendedPropertyDefinition> propSet = new List<ExtendedPropertyDefinition>();
            DoProps(ref propSet);
            cView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
            foreach (PropertyDefinitionBase pdbProp in propSet)
            {
                cView.PropertySet.Add(pdbProp);
            }

            // now go get the items. 1000 Max so must loop to get all items
            while (bMore)
            {
                findResults = fldCal.FindItems(cView);

                foreach (Item item in findResults.Items)
                {
                    cAppts.Add(item);
                }

                bMore = findResults.MoreAvailable;
                if (bMore)
                {
                    cView.Offset += iPageSize;
                }
            }

            return cAppts;
        }
    }
}
