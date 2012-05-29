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
    class Program
    {
        static void Main(string[] args)
        {

            //                                                                                                                                  
            //  UPS INTEGRATION  (beep boop beep)                                                                                                               
            //                                                                                                                                  

            Globals.Write_Gray(Environment.NewLine + "================================================");
            Globals.Write_Gray("                 UPS INTEGRATION                ");
            Globals.Write_Gray("================================================");

            Globals.Write_Green(Environment.NewLine + "================================================");
            Globals.Write_Green("               LOADING / PARSING                ");
            Globals.Write_Green("================================================");

            Soap soap = new Soap();

            XmlDocument acc = soap.getAccessRequest();
            XmlDocument qvr = soap.getQuantumViewRequest();           
            XmlDocument resp = soap.sendXmlRequest(soap.joinRequests(acc, qvr));
            //XmlDocument test = new XmlDocument();
            //test.Load("../../XML/resp4.xml");
            //XmlDocument resp = test;
            //soap.parseErrors(test);
            //soap.parseExceptions(test);
            soap.parseErrors(resp);
            soap.parseExceptions(resp);

            string testing = DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss");

            //Helpers.saveXMLDocument(qvr,"request_" + testing);
            Helpers.saveXMLDocument(resp, "response_" + testing);

            Globals.Write_Gray("===================================================" + Environment.NewLine);

            Globals.Write_Green(Environment.NewLine + "================================================");
            Globals.Write_Green("               NOTIFICATION EMAIL               ");
            Globals.Write_Green("================================================");

            Notify notify = new Notify();

            //                                                                                                                                  
            //                                                                                                                                  
            //                                                                                                                                  
                                                                                                                             
        }
    }
}
