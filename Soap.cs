using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Xml;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;


namespace UPS_QuantumViewImporter
{
    class Soap
    {
        public XmlDocument getAccessRequest()
        {
            XmlDocument outDoc = new XmlDocument();
            XmlDeclaration xmlDec = outDoc.CreateXmlDeclaration("1.0", "utf-8", null);

            try
            {
                XmlElement root = outDoc.CreateElement("AccessRequest");
                root.SetAttribute("xml:lang", "en-US");
                outDoc.AppendChild(root);

                XmlElement accessLicenseNumber = outDoc.CreateElement("AccessLicenseNumber");
                XmlText txt = outDoc.CreateTextNode(Globals.upsLicenseNum);
                accessLicenseNumber.AppendChild(txt);
                root.AppendChild(accessLicenseNumber);

                XmlElement userId = outDoc.CreateElement("UserId");
                txt = outDoc.CreateTextNode(Globals.upsUserID);
                userId.AppendChild(txt);
                root.AppendChild(userId);


                XmlElement password = outDoc.CreateElement("Password");
                txt = outDoc.CreateTextNode(Globals.upsPassword);
                password.AppendChild(txt);
                root.AppendChild(password);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Globals.errorLog("AR-01", "getAccessRequest", e.ToString());
            }

            return outDoc;
        }

        public XmlDocument getQuantumViewRequest()
        {
            XmlDocument outDoc = new XmlDocument();
            XmlDeclaration xmlDec = outDoc.CreateXmlDeclaration("1.0", "utf-8", null);

            try
            {
                XmlElement root = outDoc.CreateElement("QuantumViewRequest");
                root.SetAttribute("xml:lang", "en-US");
                outDoc.AppendChild(root);

                XmlElement request = outDoc.CreateElement("Request");
                XmlElement requestAction = outDoc.CreateElement("RequestAction");
                XmlText txt = outDoc.CreateTextNode(Globals.upsRequestAction);
                requestAction.AppendChild(txt);
                request.AppendChild(requestAction);
                root.AppendChild(request);

                XmlElement subRequest = outDoc.CreateElement("SubscriptionRequest");
                XmlElement subName = outDoc.CreateElement("Name");
                txt = outDoc.CreateTextNode(Globals.upsSubscriptionRequest);
                subName.AppendChild(txt);
                subRequest.AppendChild(subName);
                root.AppendChild(subRequest);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Globals.errorLog("QVR-01", "getQuantumViewRequest", e.ToString());
            }

            return outDoc;
        }

        public XmlDocument sendXmlRequest(string pReq)
        {
            XmlDocument docResp = null;
            string req = pReq;

            HttpWebRequest objHttpWebRequest;
            HttpWebResponse objHttpWebResponse = null;

            Stream objRequestStream = null;
            Stream objResponseStream = null;

            XmlTextReader objXMLReader;

            ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AcceptAllCertifications);

            objHttpWebRequest = (HttpWebRequest)WebRequest.Create(Globals.upsUri);

            try
            {
                byte[] bytes;
                bytes = System.Text.Encoding.ASCII.GetBytes(req);
                objHttpWebRequest.Method = "POST";
                objHttpWebRequest.ContentLength = bytes.Length;
                objHttpWebRequest.ContentType = "text/xml; encoding='utf-8'";

                objRequestStream = objHttpWebRequest.GetRequestStream();

                objRequestStream.Write(bytes, 0, bytes.Length);

                objRequestStream.Close();

                objHttpWebResponse = (HttpWebResponse)objHttpWebRequest.GetResponse();

                if (objHttpWebResponse.StatusCode == HttpStatusCode.OK)
                {
                    objResponseStream = objHttpWebResponse.GetResponseStream();

                    objXMLReader = new XmlTextReader(objResponseStream);

                    XmlDocument xmldoc = new XmlDocument();
                    xmldoc.Load(objXMLReader);

                    docResp = xmldoc;

                    objXMLReader.Close();
                }
                objHttpWebResponse.Close();
            }
            catch (WebException we)
            {
                Globals.Write_Magenta(we.ToString());
                Globals.errorLog("send-01", "sendXmlRequest", we.ToString());
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Globals.errorLog("send-02", "sendXmlRequest", e.ToString());
            }
            finally
            {
                objRequestStream.Close();
                objResponseStream.Close();
                objHttpWebResponse.Close();

                objXMLReader = null;
                objRequestStream = null;
                objResponseStream = null;
                objHttpWebRequest = null;
                objHttpWebResponse = null;
            }

            Console.ForegroundColor = ConsoleColor.Red;
            if (docResp != null)
                Console.WriteLine("Success!");
            docResp.Save(Console.Out);
            Console.ResetColor();
            
            return docResp;
        }

        public string joinRequests(XmlDocument doc1, XmlDocument doc2)
        {
            StringBuilder sb = new StringBuilder();

            string xmlHeader = "<?xml version=\"1.0\"?>";
            sb.AppendLine(xmlHeader);
            sb.AppendLine(doc1.OuterXml);
            sb.AppendLine(xmlHeader);
            sb.AppendLine(doc2.OuterXml);
            
            Globals.Write_Green(sb.ToString());

            return sb.ToString();
        }

        public void parseErrors(XmlDocument doc)
        {
            XmlNodeList errors = doc.GetElementsByTagName("Error");

            if (errors.Count > 0)
            {
                try
                {
                    for (int i = 0; i < errors.Count; i++)
                    {
                        string errCode = "";
                        string errSeverity = "";
                        string errDesc = "";
                        string errLoc = "";
                        string errLocElement = "";

                        for (int ii = 0; ii < errors[i].ChildNodes.Count; ii++)
                        {
                            string temp = errors[i].ChildNodes[ii].InnerText;
                            string node = errors[i].ChildNodes[ii].Name;

                            switch (node)
                            {
                                case "ErrorSeverity":
                                    errSeverity = temp;
                                    break;
                                case "ErrorCode":
                                    errCode = temp;
                                    break;
                                case "ErrorDescription":
                                    errDesc = temp;
                                    break;
                                case "ErrorLocation":
                                    errLoc = temp;
                                    break;
                                case "ErrorLocationElementName":
                                    errLocElement = temp;
                                    break;
                            }
                        }

                        if (errCode != "330028")
                            responseError(errCode, errDesc, errSeverity, errLoc, errLocElement);
                    }
                }
                catch (Exception e)
                {
                    Globals.Write_Magenta(e.ToString());
                    Globals.errorLog("parse-01", "parseErrors", e.ToString());
                }
            }
        }

        public void parseExceptions(XmlDocument doc)
        {            
            ArrayList hardTelling = new ArrayList();

            XmlNodeList exceptions = doc.GetElementsByTagName("Exception");

            if (exceptions.Count > 0)
            {
                try
                {
                    // Loop through each exception                                              
                    for (int i = 0; i < exceptions.Count; i++)
                    {
                        DateTime exceptionDate = DateTime.Today;
                        string exceptionReasonCode = "";
                        string exceptionDescription = "";
                        string trackingNumber = "";
                        string resolutionCode = "";
                        string resolutionDescription = "";
                        string statusCode = "";
                        string statusDescription = "";
                        int jobN = 0;
                        int shipmentNumber = 0;

                        // Loop through each child node                                         
                        for (int ii = 0; ii < exceptions[i].ChildNodes.Count; ii++)
                        {
                            string temp = exceptions[i].ChildNodes[ii].InnerText;
                            string node = exceptions[i].ChildNodes[ii].Name;

                            switch (node)
                            {
                                case "ReasonCode":
                                    exceptionReasonCode = temp;
                                    break;
                                case "ReasonDescription":
                                    exceptionDescription = temp;
                                    break;
                                case "Date":
                                    try
                                    {
                                        exceptionDate = DateTime.ParseExact(temp, "yyyyMMdd", null);
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.ToString());
                                        Globals.errorLog("parse-03", "parseExceptions", e.ToString());
                                    }
                                    break;
                                case "TrackingNumber":
                                    trackingNumber = temp;
                                    break;
                                case "PackageReferenceNumber":
                                    int hold_ship;

                                    if (Int32.TryParse(temp, out hold_ship))
                                        hardTelling.Add(hold_ship);
                                    break;
                                case "ShipmentReferenceNumber":
                                    int hold_job;
                                    if (Int32.TryParse(temp, out hold_job))
                                        jobN = hold_job;
                                    break;
                                case "Resolution":
                                    resolutionCode = exceptions[i].ChildNodes[ii].ChildNodes[0].InnerText;
                                    resolutionDescription = exceptions[i].ChildNodes[ii].ChildNodes[1].InnerText;
                                    break;
                                case "StatusCode":
                                    statusCode = temp;
                                    break;
                                case "StatusDescription":
                                    statusDescription = temp;
                                    break;
                            }
                        }

                        if (hardTelling.Count > 0)
                            shipmentNumber = Convert.ToInt32(hardTelling[hardTelling.Count - 1]);

                        if (!resolutionCode.Contains("Q1"))
                            responseException(shipmentNumber, exceptionDate, exceptionReasonCode, exceptionDescription, trackingNumber, jobN, resolutionCode, resolutionDescription, statusCode, statusDescription);
                    }
                }
                catch (Exception e)
                {
                    Globals.Write_Magenta(e.ToString());
                    Globals.errorLog("parse-02", "parseExceptions", e.ToString());
                }
            }
        }

        public void responseError(string pErrCode, string pErrDesc, string pErrSeverity, string pErrLoc, string pErrLocElement)
        {
            string errCode = pErrCode;
            string errDesc = pErrDesc;
            string errSeverity = pErrSeverity;
            string errLoc = pErrLoc;
            string errLocElement = pErrLocElement;
            int rowsInserted = 0;

            string q = @"INSERT INTO printable.dbo.CT_UPS_QVFeedErrorLog (errCode, errSeverity, errDescription, errLocation, errLocElement)
                            VALUES ('" + errCode + "','" + errSeverity + "','" + Globals.StringTool.RemoveApostrophes(errDesc) + "','" + Globals.StringTool.RemoveApostrophes(errLoc) + "','" + Globals.StringTool.RemoveApostrophes(errLocElement) + "')";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        //Globals.Write_Cyan(Environment.NewLine + q);
                        command.Connection.Open();
                        rowsInserted = (int)command.ExecuteNonQuery();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString());
                        Globals.errorLog("resp-01", "responseError", e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Globals.errorLog("resp-02", "responseError", e.ToString());
            }

            Globals.Write_Magenta(Environment.NewLine + rowsInserted.ToString() + " row inserted into error log - " + errCode);
        }

        public void responseException(int shipmentNumber, DateTime exceptionDate, string exceptionReasonCode, string exceptionDescription, string trackingNumber, int jobN, string resolutionCode, string resolutionDescription, string statusCode, string statusDescription)
        {
            int rowsInserted = 0;
            string q = "";

            if (jobN == 0)
            {
                q = @"INSERT INTO printable.dbo.CT_UPS_ExceptionLog (shipmentNumber, exceptionDate, exceptionReasonCode, exceptionDescription, trackingNumber, resolutionCode, resolutionDescription, statusCode, statusDescription)
                             VALUES ('" + shipmentNumber + "','" + exceptionDate + "','" + exceptionReasonCode + "','" + Globals.StringTool.RemoveApostrophes(exceptionDescription) + "','" + trackingNumber + "','" + resolutionCode + "','" + Globals.StringTool.RemoveApostrophes(resolutionDescription) + "','" + statusCode + "','" + Globals.StringTool.RemoveApostrophes(statusDescription) + "')";
            }
            else
            {
                q = @"INSERT INTO printable.dbo.CT_UPS_ExceptionLog (shipmentNumber, exceptionDate, exceptionReasonCode, exceptionDescription, trackingNumber, jobN, resolutionCode, resolutionDescription, statusCode, statusDescription)
                             VALUES ('" + shipmentNumber + "','" + exceptionDate + "','" + exceptionReasonCode + "','" + Globals.StringTool.RemoveApostrophes(exceptionDescription) + "','" + trackingNumber + "','" + jobN + "','" + resolutionCode + "','" + Globals.StringTool.RemoveApostrophes(resolutionDescription) + "','" + statusCode + "','" + Globals.StringTool.RemoveApostrophes(statusDescription) + "')";
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        rowsInserted = (int)command.ExecuteNonQuery();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString());
                        Globals.errorLog("resp-01", "responseException", e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Globals.errorLog("resp-02", "responseException", e.ToString());
            }

            Globals.Write_Cyan(Environment.NewLine + rowsInserted.ToString() + " row inserted into exception log (" + exceptionDescription + ")");
        }

        //HELPER METHOD THAT ACCEPTS ALL SSL CERTIFICATES FROM UPS
        public bool AcceptAllCertifications(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

    }
}
