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

//VERSION CONTROL TEST

namespace UPS_QuantumViewImporter
{
    class Helpers
    {
        public static void saveXMLDocument(XmlDocument xml, string filename)
        {
            System.IO.FileInfo file = new System.IO.FileInfo("../../XML/");
            file.Directory.Create();

            if ((filename != "") && (filename != null))
            {
                xml.Save("../../XML/" + filename + ".xml");
            }
            else
            {
                xml.Save("../../XML/" + DateTime.Now.ToString() + ".xml");
            }
        }

        public string getProdPlannerDetails(int prodPlanner)
        {
            string prodPlanner_email = "";

            string q = @"SELECT e.MailAlias
                             FROM pLogic.dbo.ProdPlanner p LEFT OUTER JOIN pLogic.dbo.Employee e
                                ON p.PlannerName = e.EmployeeName
                             WHERE ProdPlanner = " + prodPlanner;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            prodPlanner_email = reader[0].ToString();
                        }
                    }
                    catch (Exception e)
                    {
                        //errorLog("ProdPlanner-0", e.ToString(), "getProdPlannerDetails");
                        Globals.Write_Magenta(e.ToString());
                        Console.Beep();
                    }
                }
            }
            catch (Exception e)
            {
                //errorLog("ProdPlanner-1", e.ToString(), "getProdPlannerDetails");
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
            }

            if (prodPlanner_email == "")
            {
                prodPlanner_email = "GVreeman@Finelink.com";
            }

            return prodPlanner_email;
        }

        public int getProdPlanner(int job)
        {
            int prodPlanner = 0;

            string q = @"SELECT ProdPlanner
                         FROM pLogic.dbo.OpenJob
                         WHERE JobN = " + job;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        prodPlanner = Convert.ToInt32(command.ExecuteScalar());
                        //Globals.Write_Magenta(prodPlanner.ToString());
                    }
                    catch (Exception e)
                    {
                        //errorLog("getProdPlanner-0", e.ToString(), "getProdPlanner");
                        Globals.Write_Magenta(e.ToString());
                        Console.Beep();
                    }
                }
            }
            catch (Exception e)
            {
                //errorLog("getProdPlanner-1", e.ToString(), "getProdPlanner");
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
            }

            return prodPlanner;
        }

        public int getFGProdPlanner(int fgorder)
        {
            int prodPlanner = 0;

            string q = @"SELECT SalesmanN
                         FROM pLogic.dbo.FGIOrderMast
                         WHERE OrderN = " + fgorder;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        prodPlanner = Convert.ToInt32(command.ExecuteScalar());
                        //Globals.Write_Magenta(prodPlanner.ToString());
                    }
                    catch (Exception e)
                    {
                        //errorLog("getProdPlanner-0", e.ToString(), "getProdPlanner");
                        Globals.Write_Magenta(e.ToString());
                        Console.Beep();
                    }
                }
            }
            catch (Exception e)
            {
                //errorLog("getProdPlanner-1", e.ToString(), "getProdPlanner");
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
            }

            return prodPlanner;
        }

        public string getProdPlannerName(int prodPlanner)
        {
            string pp = "";
            string q = @"SELECT PlannerName
                         FROM pLogic.dbo.ProdPlanner
                         WHERE ProdPlanner = " + prodPlanner;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        pp = command.ExecuteScalar().ToString();
                        //Globals.Write_Magenta(prodPlanner.ToString());
                    }
                    catch (Exception e)
                    {
                        //errorLog("getProdPlanner-0", e.ToString(), "getProdPlanner");
                        Globals.Write_Magenta(e.ToString());
                        Console.Beep();
                    }
                }
            }
            catch (Exception e)
            {
                //errorLog("getProdPlanner-1", e.ToString(), "getProdPlanner");
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
            }

            return pp;
        }

        public DateTime getDueDate(int job)
        {
            DateTime duedate;
            duedate = System.DateTime.Today;

            string q = @"SELECT DueDate
                         FROM pLogic.dbo.OpenJob
                         WHERE JobN = " + job;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        duedate = Convert.ToDateTime(command.ExecuteScalar().ToString());
                        
                        //Globals.Write_Magenta(prodPlanner.ToString());
                    }
                    catch (Exception e)
                    {
                        //errorLog("getProdPlanner-0", e.ToString(), "getProdPlanner");
                        Globals.Write_Magenta(e.ToString());
                        Console.Beep();
                    }
                }
            }
            catch (Exception e)
            {
                //errorLog("getProdPlanner-1", e.ToString(), "getProdPlanner");
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
            }

            return duedate;
        }

        /// GET EXCEPTIONS LISTS
        public ArrayList getExceptions()
        {
            ArrayList al = new ArrayList();

            string q = @"SELECT exceptionNum
                         FROM printable.dbo.CT_UPS_ExceptionLog
                         WHERE notified = 0";
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            al.Add(Convert.ToInt32(reader[0].ToString()));
                        }
                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }

            return al;
        }

        public  int getShipmentNumber(int exceptionNum)
        {
            int shipmentnum = 0;
            string q = @"SELECT shipmentNumber
                         FROM printable.dbo.CT_UPS_ExceptionLog
                         WHERE exceptionNum = " + exceptionNum;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        shipmentnum = Convert.ToInt32(command.ExecuteScalar() ?? 0);
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                //Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                //Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }

            return shipmentnum;
        }

        public string getShippingAddress(int shipn)
        {
            string output = "";
            string q = @"SELECT ShiptoAttn, Addressee, AddrLine1, AddrLine2, AddrLine3, City, StateProv, PostalCode
                         FROM plogic.dbo.Shipments
                         WHERE ShipmentNumber = " + shipn;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();
                        
                        if (reader.HasRows)
                        {
                            output = "<ul style='list-style:none;'>";
                            while (reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    if ((reader[i].ToString().Length > 1)&&(reader[i].ToString().Trim() != ""))
                                        output += "<li>" + reader[i].ToString() + "</li>";
                                }
                            }
                            output = output + "</ul>";
                        }

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }
            return output;
        }

        public string getStatusDescription(int exceptionNum)
        {
            string statusDescription = "";
            string q = @"SELECT statusDescription
                         FROM printable.dbo.CT_UPS_ExceptionLog
                         WHERE exceptionNum = " + exceptionNum;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        statusDescription = command.ExecuteScalar().ToString();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }

            return statusDescription;
        }

        public string getTrackingNum(int exceptionNum)
        {
            string trackingnum = "";
            string q = @"SELECT trackingNumber
                         FROM printable.dbo.CT_UPS_ExceptionLog
                         WHERE exceptionNum = " + exceptionNum;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        trackingnum = command.ExecuteScalar().ToString();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                //Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                //Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }

            return trackingnum;
        }

        public string getTrackingByShip(int shipn)
        {
            
            string trackingnum = "";
            string q = @"SELECT trackingNumber
                         FROM printable.dbo.CT_UPS_ExceptionLog
                         WHERE shipmentNumber = " + shipn;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        trackingnum = command.ExecuteScalar().ToString();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        Globals.errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }

            return trackingnum;
        }

        public string getResolutionDescription(int exceptionNum)
        {
            string resolutionDescription = "";
            string q = @"SELECT resolutionDescription
                         FROM printable.dbo.CT_UPS_ExceptionLog
                         WHERE exceptionNum = " + exceptionNum;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        resolutionDescription = command.ExecuteScalar().ToString();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }

            return resolutionDescription;
        }

        public string getExceptionDescription(int exceptionNum)
        {
            string resolutionDescription = "";
            string q = @"SELECT exceptionDescription
                         FROM printable.dbo.CT_UPS_ExceptionLog
                         WHERE exceptionNum = " + exceptionNum;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        resolutionDescription = command.ExecuteScalar().ToString();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                //Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                //Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }

            return resolutionDescription;
        }

        public ArrayList getjobnumber(int shipmentnumber)
        {
            ArrayList al = new ArrayList();
            string q = @"select mainreference
                                from plogic.dbo.shipmentitems
                                where shipmentnumber = " + shipmentnumber;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            al.Add(Convert.ToInt32(reader[0].ToString()));
                        }
                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorlog("getjobs-1", e.tostring(), "get initial jobs error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("connection timeout");
                Console.Beep();
                Globals.errorLog("getjobs-te", t.ToString(), "get initial jobs timeout error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getjobs-2", e.ToString(), "get initial jobs error");
            }

            return al;

        }

        public int getfgnumber(int order)
        {
            int num = 0;
            string q = @"select FGOrder
                                from plogic.dbo.shipments
                                where shipmentnumber = " + order;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        num = Convert.ToInt32(command.ExecuteScalar());
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorlog("getjobs-1", e.tostring(), "get initial jobs error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("connection timeout");
                Console.Beep();
                Globals.errorLog("getjobs-te", t.ToString(), "get initial jobs timeout error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                //errorlog("getjobs-2", e.tostring(), "get initial jobs error");
            }

            return num;

        }

        public bool checkFG(int shipmentnum)
        {
            int fg = 0;
            string q = @"SELECT ISNULL(FGOrder,0)
                         FROM plogic.dbo.Shipments
                         WHERE ShipmentNumber = " + shipmentnum;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        fg = Convert.ToInt32(command.ExecuteScalar() ?? 0);
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //errorLog("getJobs-1", e.ToString(), "Get Initial Jobs Error");
                    }
                }
            }
            catch (TimeoutException t)
            {
                Globals.Write_Magenta("Connection Timeout");
                Console.Beep();
                Globals.errorLog("getJobs-TE", t.ToString(), "Get Initial Jobs Timeout Error");
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getJobs-2", e.ToString(), "Get Initial Jobs Error");
            }
            if (fg == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public string prepEmail(int exception_num)
        {
            int shipn = getShipmentNumber(exception_num);
            ArrayList al_jobn = new ArrayList(getjobnumber(shipn));
            string exdesc = getExceptionDescription(exception_num);
            string exres = getResolutionDescription(exception_num);
            string exstat = getStatusDescription(exception_num);
                        
            string msg = "";

            msg = "<body ><h1 style='text-align:center;'>UPS Exception Notification</h1>";
            
            msg += "<h3 style='color:red;'>Shipment Number: " + shipn + "</h3>";

            if (getTrackingNum(exception_num).Length > 0)
            {
                msg += "<h3>UPS Tracking Number: <a style= color: black;' href='http://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=";
                msg += getTrackingNum(exception_num) + "'>" + getTrackingNum(exception_num) + "  </a>";
                msg += "<span style='font-size: 11px;'> (Click for tracking) </span><br/><br/>";
            }

            if (exdesc.Length > 0)
                msg += "<b style='color:red;'>Exception Description: " + exdesc + "</b><br /><br />";
            
            if (exres.Length > 0)
                msg += "<b>Resolution Description: " + exres + "</b><br /><br /><hr>";

            if (exstat.Length > 0)
                msg += "<b>Status Description: " + exstat + "</b><br /><br /><hr>";

            msg += "<b>Shipping Address: " + getShippingAddress(shipn) + "</b><br /><hr>";

            Helpers help = new Helpers();
            al_jobn = help.RemoveDups(al_jobn);
            help = null;
            foreach (int job in al_jobn)
            {
                msg += "<b>Job Number: " + job.ToString() + "<br />";
                msg += "<b>Customer Name: " + getCustomerName(job) +  " (" + getCustomerNum(job) + ")</b><br />";
                msg += "<b>CSR: " + getProdPlannerName(getProdPlanner(job)) + "<br />";
                msg += "Due Date: " + getDueDate(job).ToString("MM/dd/yyyy") + "<br /><hr>";   
            }
            msg += "<p style ='color:#989898';>Exception Number: " + exception_num + "</p></body>";
            return msg;
                 
        }

        public string prepFGEmail(int exception_num)
        {
            int shipn = getShipmentNumber(exception_num);
            int fgnum = getfgnumber(shipn);
            string exdesc = getExceptionDescription(exception_num);
            string exres = getResolutionDescription(exception_num);
            string exstat = getStatusDescription(exception_num);
                        
            string msg = "";

            msg = "<body ><h1 style='text-align:center;'>UPS Exception Notification</h1>";
            
            msg += "<h3 style='color:red;'>Shipment Number: " + shipn + "</h3>";

            if (getTrackingNum(exception_num).Length > 0)
            {
            msg += "<h3>UPS Tracking Number: <a style= color: black;' href='http://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=";

            msg += getTrackingNum(exception_num) + "'>" + getTrackingNum(exception_num) + "  </a>";
            msg += "<span style='font-size: 11px;'> (Click for tracking) </span><br/><br/>";
            }

            if (exdesc.Length > 0)
                msg += "<b style='color:red;'>Exception Description: " + exdesc + "</b><br /><br />";
            
            if (exres.Length > 0)
                msg += "<b>Resolution Description: " + exres + "</b><br /><br /><hr>";

            if (exstat.Length > 0)
                msg += "<b>Status Description: " + exstat + "</b><br /><br /><hr>";

            msg += "<b>Shipping Address: " + getShippingAddress(shipn) + "</b><br /><hr>";      
            msg += "<b>FG Order Number: " + fgnum.ToString() + "<br />";
            msg += "<b>Customer Name: " + getFGCustomerName(getFGCustomerNum(fgnum)) +  " (" + getCustomerNum(fgnum) + ")</b><br />";
            msg += "<b>CSR: " + getProdPlannerName(getFGProdPlanner(fgnum)) + "<br /><hr>";
       
            msg += "<p style ='color:#989898';>Exception Number: " + exception_num + "</p></body>";
            return msg;
                 
        }

        public void sendEmail(string recipID, string msgBody)
        {
            ArrayList msgList = new ArrayList();

            msgList.Clear();
            msgList.Add(recipID);
                

                foreach (string item in msgList)
                {
                    try
                    {
                        MailMessage message = new MailMessage();
                        message.To.Add(item);
                        message.Subject = "UPS Exception Notification";
                        message.From = new MailAddress("UPSExceptions@Finelink.com");
                        message.Body = msgBody;
                        message.ReplyTo = new MailAddress("GVreeman@Finelink.com");
                        message.IsBodyHtml = true;
                        System.Net.Mail.SmtpClient smtp = Globals.get_smtpClient;
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.Write("Sending email to " + message.To.ToString());
                        smtp.Send(message);
                        Console.WriteLine(" - Success");
                        Console.ResetColor();
                        message.Dispose();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("SendEmail Error - RecipID:" + recipID + " - Receiver:" + item + " - Msg:" + msgBody);
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        Globals.errorLog("Email-1", e.ToString(), "SendEmail Error - RecipID:" + recipID + " - Receiver:" + item + " - Msg:" + msgBody);
                    }
                }         
        }

        public int getCustomerNum(int pJobN)
        {
            // Retrieve from pLogic.dbo.OpenJob = CustomerN
            int CustomerN = 0;
            string q_getCust = @"SELECT TOP 1 CustomerN
                                 FROM pLogic.dbo.OpenJob
                                 WHERE JobN = " + pJobN;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q_getCust, conn);
                    try
                    {
                        command.Connection.Open();
                        CustomerN = Convert.ToInt32(command.ExecuteScalar());

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        Globals.errorLog("getCust-1", e.ToString(), pJobN + " - Job - Get CustomerN Error");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return CustomerN;
        }

        public int getFGCustomerNum(int order)
        {
            // Retrieve from pLogic.dbo.OpenJob = CustomerN
            int CustomerN = 0;
            string q_getCust = @"SELECT TOP 1 CustomerN
                                 FROM pLogic.dbo.FGIOrderMast
                                 WHERE OrderN = " + order;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q_getCust, conn);
                    try
                    {
                        command.Connection.Open();
                        CustomerN = Convert.ToInt32(command.ExecuteScalar());

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                        //Globals.errorLog("getCust-1", e.ToString(), pJobN + " - Job - Get CustomerN Error");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return CustomerN;
        }

        public string getCustomerName(int pJobN)
        {
            string temp = "";
            string q = @"SELECT TOP 1 CustomerName
                         FROM pLogic.dbo.Customer
                         WHERE CustomerN = " + getCustomerNum(pJobN);

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        temp = command.ExecuteScalar().ToString();

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString() + " - " + pJobN);
                        Console.Beep();
                        Globals.errorLog("getCustomerName-1", e.ToString(), "Get CustomerName Error");
                    }
                }
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getCustomerName-0", e.ToString(), "Get CustomerName Error"); ;
            }
            return temp;
        }

        public string getFGCustomerName(int custnum)
        {
            string temp = "";
            string q = @"SELECT TOP 1 CustomerName
                         FROM pLogic.dbo.Customer
                         WHERE CustomerN = " + custnum;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        temp = command.ExecuteScalar().ToString();

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString() + " - " + custnum);
                        Console.Beep();
                        Globals.errorLog("getCustomerName-1", e.ToString(), "Get CustomerName Error");
                    }
                }
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
                Console.Beep();
                Globals.errorLog("getCustomerName-0", e.ToString(), "Get CustomerName Error"); ;
            }
            return temp;
        }

        public ArrayList RemoveDups(ArrayList items)
        {
            ArrayList noDups = new ArrayList();

            foreach (int strItem in items)
            {
                if (!noDups.Contains(strItem))
                {
                    noDups.Add(strItem);
                }
            }
            noDups.Sort();
            return noDups;
        }

        public ArrayList RemoveDupString(ArrayList items)
        {
            ArrayList noDups = new ArrayList();

            foreach (string strItem in items)
            {
                if (!noDups.Contains(strItem))
                {
                    noDups.Add(strItem);
                }
            }
            noDups.Sort();
            return noDups;
        }
    }
}

    