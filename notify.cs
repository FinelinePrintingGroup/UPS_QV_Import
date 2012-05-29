using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.SqlTypes;

namespace UPS_QuantumViewImporter
{
    class Notify
    {
        public Notify()
        {
            reporting();
            rayPort();
        }

        public void reporting()
        {
            Helpers help = new Helpers();
            Globals globals = new Globals();
            ArrayList al_exceptions = new ArrayList(help.getExceptions());
            ArrayList al_jerbs = new ArrayList();
            //ArrayList al_fgjerbs = new ArrayList();
            ArrayList al_prod = new ArrayList();
            ArrayList al_prodetails = new ArrayList();
            int shipnum;
            string email;

            foreach (int exnum in al_exceptions)
            {
                Console.WriteLine(exnum);

                try
                {
                    shipnum = help.getShipmentNumber(exnum);
                    Console.WriteLine(shipnum);
                    Console.WriteLine(help.checkFG(shipnum));
                    
                    // FG ORDER SHIPMENT                                                                    
                    if (help.checkFG(shipnum))
                    {
                        int fgnum = help.getfgnumber(shipnum);
                        string fgprodetails = help.getProdPlannerDetails(help.getFGProdPlanner(fgnum));
                        string mail = help.prepFGEmail(exnum);
                        help.sendEmail("keenan@finelink.com", mail);
                        help.sendEmail("shawnh@finelink.com", mail);
                    }

                    // JOBS SHIPMENT                                                                        
                    else
                    {
                        // GET ALL THE JOBS ON THE SHIPMENT                             
                        al_jerbs.Clear();
                        al_jerbs = help.getjobnumber(shipnum);

                        al_prod.Clear();
                        al_prodetails.Clear();

                        //Console.WriteLine("Not a FG Order");
                        foreach (int jerb in al_jerbs)
                        { 
                            al_prod.Add(Convert.ToInt32(help.getProdPlanner(jerb)));
                            
                            foreach (int prod in al_prod)
                            {
                                al_prodetails.Add(help.getProdPlannerDetails(prod));
                            }
                        }

                        al_prodetails = help.RemoveDupString(al_prodetails);

                        email = help.prepEmail(exnum);
                       
                        foreach (string s in al_prodetails)
                        { 
                            //Console.WriteLine(s);
                            help.sendEmail("keenan@finelink.com", email);
                            help.sendEmail("shawnh@finelink.com", email);
                            //help.sendEmail("gvreeman@finelink.com", email);
                        }
                    }

                    
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }               
               
                //jobnum = help.getJobNumber(shipnum);

            }

        }

        public void rayPort()
        {
            Helpers help = new Helpers();
            ArrayList al_exceptions = new ArrayList(help.getExceptions());

            string msg = "";

            msg = "<body><h1 style='text-align:center;'>UPS Exceptions Report</h1>";
            msg += "<table style='border: 1px solid black; text-align: left; border-collapse:collapse'>";
            msg += "<tr style='border: 1px solid black; background-color: #c0c0c0; font-weight: bold;'><td style='text-align:center; border-bottom: 1px solid black;'>SHIPMENT NUMBER</td><td td style='text-align:center; border-bottom: 1px solid black;'>EXCEPTION DESCRIPTION</td>";
            msg += "<td style='text-align:center; border-bottom: 1px solid black;'>TRACKING NUMBER</td></tr>";

            foreach (int ex in al_exceptions)
            {
                string exDesc = help.getExceptionDescription(ex);
                string reDesc = help.getResolutionDescription(ex);
                string stDesc = help.getStatusDescription(ex);

                msg += "<tr style='border-bottom: 1px solid black;'><td style='text-align:center; border-bottom: 1px solid black; border-right: 1px solid black;' height='45'>";
                msg += help.getShipmentNumber(ex) + "</td><td style=' border-bottom: 1px solid black; border-right: 1px solid black;' height='45'>";

                if (exDesc.Trim().Length > 5)
                    msg += exDesc + "<br /><br />";

                if (reDesc.Trim().Length > 5)
                    msg += reDesc + "<br /><br />";

                if (stDesc.Trim().Length > 5)
                    msg += stDesc + "<br /><br />";

                msg += "</td><td style='text-align:center; border-bottom: 1px solid black; border-right: 1px solid black;' height='45'>";
                msg += "<a style= color: black;' href='http://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=";
                msg += help.getTrackingNum(ex) + "'>" + help.getTrackingNum(ex) + "  </a></td></tr>";
                flipDaBitYooooooo(ex);
            }

            msg += "</table>";

            if (al_exceptions.Count > 0)
            {
                help.sendEmail("rayv@finelink.com", msg);
                help.sendEmail("keenan@finelink.com", msg);
                help.sendEmail("shawnh@finelink.com", msg);
            }
            help = null;
        }

        protected void flipDaBitYooooooo(int exceptionNum)
        {
            int ShutUp = 0;

            string q = @"UPDATE printable.dbo.CT_UPS_ExceptionLog
                         SET notified = 1
                         WHERE exceptionNum = " + exceptionNum;
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        conn.Open();
                        ShutUp = Convert.ToInt32(command.ExecuteScalar() ?? 0);
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Globals.Write_Magenta(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Globals.Write_Magenta(e.ToString());
            }
        }
    }
}
