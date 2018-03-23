﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using static CalVerifier.Globals;
using static CalVerifier.Utils;

namespace CalVerifier
{
    public class Process
    {
        // properties that are used for the tests
        public static string strSubject = "";                 //PR_SUBJECT
        public static string strOrganizerName = "";           //PR_SENT_REPRESENTING_NAME_W
        public static string strOrganizerAddr = "";           //PR_SENT_REPRESENTING_EMAIL_ADDRESS_W 
        public static string strOrganizerAddrType = "";       //PR_SENT_REPRESENTING_ADDRTYPE_W
        public static string strSenderName = "";              //PR_SENDER_NAME_W
        public static string strSenderAddr = "";              //PR_SENDER_EMAIL_ADDRESS_W
        public static string strMsgClass = "";                //PR_MESSAGE_CLASS
        public static string strLastModified = "";            //PR_LAST_MODIFICATION_TIME
        public static string strLastModifiedBy = "";          //PR_LAST_MODIFIER_NAME_W
        public static string strEntryID = "";                 //PR_ENTRYID
        public static string strMsgSize = "";                 //PR_MESSAGE_SIZE 
        public static string strDeliveryTime = "";            //PR_MESSAGE_DELIVERY_TIME
        public static string strHasAttach = "";               //PR_HASATTACH
        public static string strMsgStatus = "";               //PR_MSG_STATUS
        public static string strCreateTime = "";              //PR_CREATION_TIME
        public static string strRecurring = "";               //dispidRecurring
        public static string strRecurType = "";               //dispidRecurType
        public static string strStartWhole = "";              //dispidApptStartWhole
        public static string strEndWhole = "";                //dispidApptEndWhole
        public static string strApptStateFlags = "";          //dispidApptStateFlags
        public static string strLocation = "";                //dispidLocation
        public static string strTZDesc = "";                  //dispidTimeZoneDesc
        public static string strAllDay = "";                  //dispidApptSubType
        public static string strRecurBlob = "";               //dispidApptRecur
        public static string strIsRecurring = "";             //PidLidIsRecurring
        public static string strGlobalObjID = "";             //PidLidGlobalObjectId
        public static string strCleanGlobalObjID = "";        //PidLidCleanGlobalObjectId
        public static string strAuxFlags = "";                //dispidApptAuxFlags
        public static string strIsException = "";             //PidLidIsException
        public static string strTZStruct = "";                //dispidTimeZoneStruct
        public static string strTZDefStart = "";              //dispidApptTZDefStartDisplay
        public static string strTZDefEnd = "";                //dispidApptTZDefEndDisplay
        public static string strTZDefRecur = "";              //dispidApptTZDefRecur
        public static string strPropDefStream = "";           //dispidPropDefStream

        // Test this Calendar Item's properties.
        public static void ProcessItem(Appointment appt)
        {
            string strLogItem = "Problem item: " + strSubject + ", " + strLocation + ", " + strStartWhole + ", " + strEndWhole;
            List<string> strErrors = new List<string>();
            bool bErr = false;
            bool bWarn = false;
            
            foreach (string strVal in appt.Categories)
            {
                if (strVal.ToUpper() == "HOLIDAY")
                {
                    return; // we will skip testing holiday items since they are imported and should be okay
                }
            }
            
            // populate the values for the properties
            GetPropsReadable(appt);

            //get other types of values as needed from the string values
            DateTime dtStart = DateTime.Parse(strStartWhole);
            DateTime dtEnd = DateTime.Parse(strEndWhole);
            NameResolutionCollection ncCol = Utils.exService.ResolveName(strOrganizerAddr);
            string strOrganizerSMTP = "";
            if (ncCol.Count > 0 && !string.IsNullOrEmpty(ncCol[0].Mailbox.Address))
            {
                strOrganizerSMTP = ncCol[0].Mailbox.Address;
            }
            else
            {
                strOrganizerSMTP = strOrganizerAddr;
            }

            // really actually start testing props
            if (string.IsNullOrEmpty(strSubject))
            {
                bWarn = true;
                strErrors.Add("   WARNING: Subject is empty/missing.");
                iWarn++;
            }
            if (string.IsNullOrEmpty(strDeliveryTime))
            {
                bErr = true;
                strErrors.Add("   ERROR: Missing required Delivery Time property.");
                iErrors++;
            }

            if (string.IsNullOrEmpty(strRecurring))
            {
                bErr = true;
                strErrors.Add("   ERROR: Missing required Recurring property.");
                iErrors++;
            }
            else
            {
                if (strRecurring.ToUpper() == "TRUE")
                {
                    iRecurItems++;
                    if (iRecurItems == 1299)
                    {
                        bErr = true;
                        strErrors.Add("   ERROR: Reached limit of 1300 Recurring Appointments. Delete some older recurring appointments to correct this.");
                        iErrors++;
                    }
                    if (iRecurItems == 1250)
                    {
                        bWarn = true;
                        strErrors.Add("   WARNING: Approaching limit of 1300 Recurring Appointments. Delete some older recurring appointments to correct this.");
                        iWarn++;
                    }
                }
            }

            if (string.IsNullOrEmpty(strStartWhole))
            {
                bErr = true;
                strErrors.Add("   ERROR: Missing required Start Time property.");
                iErrors++;
            }
            else // not empty/missing, but might still have problems
            {
                if (dtEnd <= dtStart)
                {
                    bErr = true;
                    strErrors.Add("   ERROR: Start Time is greater than or equal to End Time.");
                    iErrors++;
                }

                if (TimeCheck(dtStart))  
                {
                    bErr = true;
                    strErrors.Add("   ERROR: Start Time is not set correctly."); 
                    iErrors++;
                }
            }

            if (string.IsNullOrEmpty(strEndWhole))
            {
                bErr = true;
                strErrors.Add("   ERROR: Missing required End Time property.");
                iErrors++;
            }
            else // not empty/missing, but might still have problems
            {

                if (TimeCheck(dtEnd))
                {
                    bErr = true;
                    strErrors.Add("   ERROR: End Time is not set correctly.");
                    iErrors++;
                }
            }

            if (string.IsNullOrEmpty(strOrganizerAddr))
            {
                if (int.Parse(strApptStateFlags) > 0) // if no Organizer Address AND this is a meeting then that's bad.
                {
                    bErr = true;
                    strErrors.Add("   ERROR: Missing required Organizer Address property.");
                    iErrors++;
                }
            }

            if (string.IsNullOrEmpty(strSenderName))
            {
                if (int.Parse(strApptStateFlags) > 0) // if no Sender Name AND this is a meeting then that's bad.
                {
                    bErr = true;
                    strErrors.Add("   ERROR: Missing required Sender Name property.");
                    iErrors++;
                }
            }

            if (string.IsNullOrEmpty(strSenderAddr))
            {
                if (int.Parse(strApptStateFlags) > 0) // if no Sender Address AND this is a meeting then that's bad.
                {
                    bErr = true;
                    strErrors.Add("   ERROR: Missing required Sender Address property.");
                    iErrors++;
                }
            }

            if (string.IsNullOrEmpty(strMsgClass))
            {
                bErr = true;
                strErrors.Add("   ERROR: Missing required Message Class property.");
                iErrors++;
            }
            else
            {
                bool bFound = false;
                foreach (string strClass in calMsgClasses)
                {
                    if (strClass == strMsgClass)
                    {
                        bFound = true;
                        break; // if one of the known classes then all is good.
                    }
                }

                if (!bFound)
                {
                    bWarn = true;
                    strErrors.Add("   WARNING: Unknown or incorrect Message Class " + strMsgClass + " is set on this item.");
                    iWarn++;
                }
            }

            if (!(string.IsNullOrEmpty(strMsgSize)))
            {
                int iSize = int.Parse(strMsgSize);
                string strNum = "";
                

                if (iSize >= 52428800)
                {
                    strNum = "50M";
                }
                else if (iSize >= 26214400)
                {
                    strNum = "25M";
                }
                else if (iSize >= 10485760)
                {
                    strNum = "10M";
                }

                if (iSize >= 10485760) // if >= 10M then one of the above is true...
                {
                    bWarn = true;
                    iWarn++;
                    if (strHasAttach.ToUpper() == "TRUE" && strRecurring.ToUpper() == "TRUE")
                    {
                        strErrors.Add("   WARNING: Message size exceeds " + strNum + " which may indicate a problematic long-running recurring meeting.");
                    }
                    else if (strHasAttach.ToUpper() == "TRUE")
                    {
                        strErrors.Add("   WARNING: Message size exceeds " + strNum + " but is not set as recurring. Might have large and/or many attachments.");
                    }
                    else
                    {
                        strErrors.Add("   WARNING: Message size exceeds " + strNum + " but has no attachments. Might have some large problem properties.");
                    }
                }

            }

            if (string.IsNullOrEmpty(strApptStateFlags)) //
            {
                bErr = true;
                strErrors.Add("   ERROR: Missing required Appointment State property.");
                iErrors++;
            }
            else
            {
                // check for meeting hijack items
                switch (strApptStateFlags)
                {
                    case "0": // Non-meeting appointment
                        {
                            //single appointment I made in my Calendar
                            break;
                        }
                    case "1": // Meeting and I am the Organizer
                        {
                            if (!string.IsNullOrEmpty(strOrganizerAddr) && !string.IsNullOrEmpty(strOrganizerSMTP))
                            {
                                if (!(strOrganizerSMTP.ToUpper() == strSMTPAddr)) // this user's email should match with the Organizer. If not then error.
                                {
                                    bErr = true;
                                    strErrors.Add("   ERROR: Organizer properties are in conflict.");
                                    strErrors.Add("   Organizer Address: " + strOrganizerAddr);
                                    strErrors.Add("   Appt State: " + strDisplayName + " is the Organizer");
                                    iErrors++;
                                }
                            }
                            break;
                        }
                    case "2": // Received item - shouldn't be in this state
                        {
                            bErr = true;
                            strErrors.Add("   ERROR: Appointment State is an incorrect value.");
                            iErrors++;
                            break;
                        }
                    case "3": // Meeting item that I received - I am an Attendee
                        {
                            if (!string.IsNullOrEmpty(strOrganizerAddr) && !string.IsNullOrEmpty(strOrganizerSMTP))
                            {
                                if (strOrganizerSMTP.ToUpper() == strSMTPAddr) // this user's email should NOT match with the Organizer. If it does then error.
                                {
                                    bErr = true;
                                    strErrors.Add("   ERROR: Organizer properties are in conflict.");
                                    strErrors.Add("   Organizer Address: " + strOrganizerAddr);
                                    strErrors.Add("   Appt State: " + strDisplayName + " is an Attendee");
                                    iErrors++;
                                }
                            }
                            break;
                        }
                    default: // nothing else matters yet - can add later if needed
                        {
                            break;
                        }
                }
            }

            if (string.IsNullOrEmpty(strTZDefStart))
            {
                if (strRecurring.ToUpper() == "TRUE")
                {
                    bErr = true;
                    strErrors.Add("   ERROR: Missing required Timezone property.");
                    iErrors++;
                }
            }

            if (bErr || bWarn)
            {
                outLog.WriteLine(strLogItem);
                // AND log out each line in the List of errors
                foreach (string strLine in strErrors)
                {
                    outLog.WriteLine(strLine);
                }
            } 
        }

        // Populate the property values for each of the props the app checks on.
        // Some tests require multiple props, so best to go ahead and just get them all first.
        public static void GetPropsReadable(Appointment appt)
        {
            string strHexTag = "";
            string strPropName = "";
            string strSetID = "";
            string strGUID = "";
            string strValue = "";
            string strType = "";
            string strKeywords;

            foreach (ExtendedProperty extProp in appt.ExtendedProperties)
            {
                // Get the Tag
                if (extProp.PropertyDefinition.Tag.HasValue)
                {
                    strHexTag = extProp.PropertyDefinition.Tag.Value.ToString("X4");
                }
                else if (extProp.PropertyDefinition.Id.HasValue)
                {
                    strHexTag = extProp.PropertyDefinition.Id.Value.ToString("X4");
                }

                // Get the SetID for named props
                if (extProp.PropertyDefinition.PropertySetId.HasValue)
                {
                    strGUID = extProp.PropertyDefinition.PropertySetId.Value.ToString("B");
                    strSetID = PropSet.GetSetIDFromGUID(strGUID);
                }

                // Get the Property Type
                strType = extProp.PropertyDefinition.MapiType.ToString();

                // Get the Prop Name
                strPropName = PropSet.GetPropNameFromTag(strHexTag, strSetID);

                // if it's binary then convert it to a string-ized binary - will be converted using MrMapi
                if (strType == "Binary")
                {
                    byte[] binData = extProp.Value as byte[];
                    strValue = GetStringFromBytes(binData);
                }
                else
                {
                    if (extProp.Value != null)
                    {
                        strValue = extProp.Value.ToString();
                    }
                }

                switch (strPropName)
                {
                    case "PR_SUBJECT_W":
                        {
                            strSubject = strValue;
                            break;
                        }
                    case "PR_SENT_REPRESENTING_NAME_W":
                        {
                            strOrganizerName = strValue;
                            break;
                        }
                    case "PR_SENT_REPRESENTING_EMAIL_ADDRESS_W":
                        {
                            strOrganizerAddr = strValue;
                            break;
                        }
                    case "PR_SENT_REPRESENTING_ADDRTYPE_W":
                        {
                            strOrganizerAddrType = strValue;
                            break;
                        }
                    case "PR_SENDER_NAME_W":
                        {
                            strSenderName = strValue;
                            break;
                        }
                    case "PR_SENDER_EMAIL_ADDRESS_W":
                        {
                            strSenderAddr = strValue;
                            break;
                        }
                    case "PR_MESSAGE_CLASS":
                        {
                            strMsgClass = strValue;
                            break;
                        }
                    case "PR_LAST_MODIFICATION_TIME":
                        {
                            strLastModified = strValue;
                            break;
                        }
                    case "PR_LAST_MODIFIER_NAME_W":
                        {
                            strLastModifiedBy = strValue;
                            break;
                        }
                    case "PR_ENTRYID":
                        {
                            strEntryID = strValue;
                            break;
                        }
                    case "PR_MESSAGE_SIZE":
                        {
                            strMsgSize = strValue;
                            break;
                        }
                    case "PR_MESSAGE_DELIVERY_TIME":
                        {
                            strDeliveryTime = strValue;
                            break;
                        }
                    case "PR_HASATTACH":
                        {
                            strHasAttach = strValue;
                            break;
                        }
                    case "PR_MSG_STATUS":
                        {
                            strMsgStatus = strValue;
                            break;
                        }
                    case "PR_CREATION_TIME":
                        {
                            strCreateTime = strValue;
                            break;
                        }
                    case "dispidRecurring":
                        {
                            strRecurring = strValue;
                            break;
                        }
                    case "dispidRecurType":
                        {
                            strRecurType = strValue;
                            break;
                        }
                    case "dispidApptStartWhole":
                        {
                            strStartWhole = strValue;
                            break;
                        }
                    case "dispidApptEndWhole":
                        {
                            strEndWhole = strValue;
                            break;
                        }
                    case "dispidApptStateFlags":
                        {
                            strApptStateFlags = strValue;
                            break;
                        }
                    case "dispidLocation":
                        {
                            strLocation = strValue;
                            break;
                        }
                    case "dispidTimeZoneDesc":
                        {
                            strTZDesc = strValue;
                            break;
                        }
                    case "dispidApptSubType":
                        {
                            strAllDay = strValue;
                            break;
                        }
                    case "dispidApptRecur":
                        {
                            strRecurBlob = strValue;
                            break;
                        }
                    case "PidLidIsRecurring":
                        {
                            strIsRecurring = strValue;
                            break;
                        }
                    case "PidLidGlobalObjectId":
                        {
                            strGlobalObjID = strValue;
                            break;
                        }
                    case "PidLidCleanGlobalObjectId":
                        {
                            strCleanGlobalObjID = strValue;
                            break;
                        }
                    case "dispidApptAuxFlags":
                        {
                            strAuxFlags = strValue;
                            break;
                        }
                    case "PidLidIsException":
                        {
                            strIsException = strValue;
                            break;
                        }
                    case "Keywords":
                        {
                            strKeywords = strValue;
                            break;
                        }
                    case "dispidTimeZoneStruct":
                        {
                            strTZStruct = strValue;
                            break;
                        }
                    case "dispidApptTZDefStartDisplay":
                        {
                            strTZDefStart = strValue;
                            break;
                        }
                    case "dispidApptTZDefEndDisplay":
                        {
                            strTZDefEnd = strValue;
                            break;
                        }
                    case "dispidApptTZDefRecur":
                        {
                            strTZDefRecur = strValue;
                            break;
                        }
                    case "dispidPropDefStream":
                        {
                            strPropDefStream = strValue;
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }
            return;
        }

        // EWS does not return a string-ized hex blob, and need it for MrMapi conversion
        public static string GetStringFromBytes(byte[] bytes)
        {
            StringBuilder ret = new StringBuilder();
            foreach (byte b in bytes)
            {
                ret.Append(Convert.ToString(b, 16).PadLeft(2, '0'));
            }

            return ret.ToString().ToUpper();
        }
    }
}