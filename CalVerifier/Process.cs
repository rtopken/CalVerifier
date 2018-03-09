using System;
using System.IO;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

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
        public static string strKeywords = "";                //Keywords
        public static string strTZStruct = "";                //dispidTimeZoneStruct
        public static string strTZDefStart = "";              //dispidApptTZDefStartDisplay
        public static string strTZDefEnd = "";                //dispidApptTZDefEndDisplay
        public static string strTZDefRecur = "";              //dispidApptTZDefRecur
        public static string strPropDefStream = "";           //dispidPropDefStream

        // Test this Calendar Item's properties.
        public static void ProcessItem(Appointment appt)
        {
            // populate the values for the properties
            GetPropsReadable(appt);

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
                if (strType == "StringArray")
                {
                    strPropName = "Keywords"; // this is the only string array prop-value this tool consumes
                }
                else
                {
                    strPropName = PropSet.GetPropNameFromTag(strHexTag, strSetID);
                }

                // if it's binary then convert it to a string-ized binary - will be converted using MrMapi
                if (strType == "Binary")
                {
                    byte[] binData = extProp.Value as byte[];
                    strValue = GetStringFromBytes(binData);
                }
                else if (strType == "StringArray")
                {
                    strKeywords = extProp.Value.ToString();
                }
                else
                {
                    strValue = extProp.Value.ToString();
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

        // EWS does nto return a string-ized hex blob, and need it for MrMapi conversion
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