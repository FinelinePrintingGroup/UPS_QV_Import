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
using System.Globalization;
using System.Text.RegularExpressions;

namespace UPS_QuantumViewImporter
{
    class Globals
    {
        /// UPS Quantum View Information                                                                                                                 
                                                                                                                                                      
        //public static string upsLicenseNum = "0C85F7E0B209F0D8";
        public static string upsLicenseNum = "6C9898219F056698";
        public static string upsUserID = "gvreeman";
        public static string upsPassword = "M$uzik9888";

        public static string upsRequestAction = "QVEvents";
        public static string upsSubscriptionRequest = "FinelineOut";

        //public static Uri upsUri = new Uri(@"https://wwwcie.ups.com/ups.app/xml/QVEvents"); //Testing
        public static Uri upsUri = new Uri(@"https://www.ups.com/ups.app/xml/QVEvents");  //Production
        
        
        /// Logic Connections                                                                                                                                                                                                                                                                     
        //public static string logicConnString = "Data Source=SQL1; Initial Catalog=devLogic; User ID=FPGwebservice; Password=kissmygrits";
        public static string printableConnString = "Data Source=SQL1; Initial Catalog=printable; User ID=FPGwebservice; Password=kissmygrits";
        public static string logicConnString = "Data Source=SQL1;Initial Catalog=pLogic;User ID=FPGwebservice;Password=kissmygrits";
        //public static string printableConnString = "Data Source=SQL1;Initial Catalog=printable;User ID=FPGwebservice;Password=kissmygrits";
        

        
        /// Fed-Ex Information                                                                                                                                                       
        public static string fedEx = null;

                                                                                                                                             

        /// SMTP Connection Settings
        private static SmtpClient smtpClient = new SmtpClient("192.168.240.27");

        public static SmtpClient get_smtpClient
        {
            get { return Globals.smtpClient; }
        }


        /// SQL Server DateFormat
        private static string sqlDateFormat = "yyyy-MM-dd HH:mm:ss.fff";

        public static string get_sqlDateFormat
        {
            get { return Globals.sqlDateFormat; }
        }

        private static string sqlShortDateFormat = "yyyy-MM-dd";

        public static string get_sqlShortDateFormat
        {
            get { return Globals.sqlShortDateFormat; }
        }


        /// Helper Tools
        public static class StringTool
        {
            /// <summary>
            /// Basic Truncate function that takes the source string and truncates it to the length specified. Verifies that the string is greater than the length.
            /// </summary>
            /// <param name="source">Source string to be truncated</param>
            /// <param name="length">Number of characters after which to truncate the source string</param>
            /// <returns>Returns the truncated source string</returns>
            public static string Truncate(string source, int length)
            {
                if (source.Length > length)
                {
                    source = source.Substring(0, length);
                }
                return source;
            }

            /// <summary>
            /// Basic Truncate function that takes the source string and truncates it to the length specified.
            /// </summary>
            /// <param name="source">Source string to be truncated</param>
            /// <param name="length">Number of characters after which to truncate the source string</param>
            /// <returns>Returns the truncated source string</returns>
            public static string Truncate2(string source, int length)
            {
                return source.Substring(0, Math.Min(length, source.Length));
            }

            /// <summary>
            /// Escapes all single quotes in the source string with an additional single quote so that it can be used in a SQL statement.
            /// </summary>
            /// <param name="source">Source string possible containing single quotes</param>
            /// <returns>Source string with all single quotes escaped for use in SQL statement</returns>
            public static string RemoveApostrophes(string source)
            {
                return source.Replace("'", "''");
            }

            /// <summary>
            /// Takes a string and splits it into multiple strings of max allowedLength.
            /// </summary>
            /// <param name="str">The string</param>
            /// <param name="allowedLength">How long the string can be</param>
            /// <returns>A list array of strings</returns>
            public static List<string> splitStrings(string str, int allowedLength)
            {
                // Trim leading and trailing whitespace 
                str = str.Trim();

                // Remove newlines and carriage returns 
                str = str.Replace(Environment.NewLine, " ");
                str = str.Replace("\n", "X");
                str = str.Replace("\r", "X");

                // Remove excessive spacing             
                str = str.Replace("   ", " ");
                str = str.Replace("  ", " ");

                var ret1 = str.Split(' ');
                var ret2 = new List<string>();
                ret2.Add("");
                int index = 0;
                foreach (var item in ret1)
                {
                    if (item.Length + 1 + ret2[index].Length <= allowedLength)
                    {
                        ret2[index] += ' ' + item;
                        if (ret2[index].Length >= allowedLength)
                        {
                            ret2.Add("");
                            index++;
                        }
                    }
                    else
                    {
                        ret2.Add(item);
                        index++;
                    }
                }
                return ret2;
            }
        }

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
                xml.Save("../../XML/" + DateTime.Now.ToShortDateString() + ".xml");
            }
        }

        public static void ReadWriteStream(Stream readStream, Stream writeStream)
        {
            int Length = 256;
            Byte[] buffer = new Byte[Length];
            int bytesRead = readStream.Read(buffer, 0, Length);
            // write the required bytes
            while (bytesRead > 0)
            {
                writeStream.Write(buffer, 0, bytesRead);
                bytesRead = readStream.Read(buffer, 0, Length);
            }
            readStream.Close();
            writeStream.Close();
        }

        public static void pointlessFinelineASCII()
        {
            Write_Cyan(@"
                                 ___           ___           ___                    ___           ___     
      ___            ___        /  /\         /  /\         /  /\       ___        /  /\         /  /\    
     /  /\          /__/\      /  /::|       /  /::\       /  /:/      /__/\      /  /::|       /  /::\   
    /  /::\         \__\:\    /  /:|:|      /  /:/\:\     /  /:/       \__\:\    /  /:|:|      /  /:/\:\  
   /  /:/\:\        /  /::\  /  /:/|:|__   /  /::\ \:\   /  /:/        /  /::\  /  /:/|:|__   /  /::\ \:\ 
  /  /::\ \:\    __/  /:/\/ /__/:/ |:| /\ /__/:/\:\ \:\ /__/:/      __/  /:/\/ /__/:/ |:| /\ /__/:/\:\ \:\
 /__/:/\:\ \:\  /__/\/:/    \__\/  |:|/:/ \  \:\ \:\_\/ \  \:\     /__/\/:/    \__\/  |:|/:/ \  \:\ \:\_\/
 \__\/  \:\_\/  \  \::/         |  |:/:/   \  \:\ \:\    \  \:\    \  \::/         |  |:/:/   \  \:\ \:\  
      \  \:\     \  \:\         |__|::/     \  \:\_\/     \  \:\    \  \:\         |__|::/     \  \:\_\/  
       \__\/      \__\/         /__/:/       \  \:\        \  \:\    \__\/         /__/:/       \  \:\    
                                \__\/         \__\/         \__\/                  \__\/         \__\/   " + Environment.NewLine + Environment.NewLine);
        }

        public static void pointlessFinelineDevASCII()
        {
            Write_Green(@"  
  _____ ___ _   _ _____ _     ___ _   _ _____   ____  _______     __
 |  ___|_ _| \ | | ____| |   |_ _| \ | | ____| |  _ \| ____\ \   / /
 | |_   | ||  \| |  _| | |    | ||  \| |  _|   | | | |  _|  \ \ / / 
 |  _|  | || |\  | |___| |___ | || |\  | |___  | |_| | |___  \ V /  
 |_|   |___|_| \_|_____|_____|___|_| \_|_____| |____/|_____|  \_/   
                                                                    ");
        }

        public class RegexUtilities
        {
            bool invalid = false;

            public bool IsValidEmail(string strIn)
            {
                invalid = false;
                if (String.IsNullOrEmpty(strIn))
                    return false;

                // Use IdnMapping class to convert Unicode domain names.
                strIn = Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper);
                if (invalid)
                    return false;

                // Return true if strIn is in valid e-mail format.
                return Regex.IsMatch(strIn,
                       @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                       @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
                       RegexOptions.IgnoreCase);
            }

            private string DomainMapper(Match match)
            {
                // IdnMapping class with default property values.
                IdnMapping idn = new IdnMapping();

                string domainName = match.Groups[2].Value;
                try
                {
                    domainName = idn.GetAscii(domainName);
                }
                catch (ArgumentException)
                {
                    invalid = true;
                }
                return match.Groups[1].Value + domainName;
            }
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

        /// Console                
        public static void Write_Red(string value)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_White(string value)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_Cyan(string value)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_Yellow(string value)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_Green(string value)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_Magenta(string value)
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_Gray(string value)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_DarkGray(string value)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(value);
            Console.ResetColor();
        }
        public static void Write_Highlight(string value)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.BackgroundColor = ConsoleColor.White;
            Console.WriteLine(value);
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ResetColor();
        }


        /// XML -> Console
        public static void ConsOut_Red(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_Blue(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_Cyan(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_DarkBlue(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_DarkCyan(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_DarkGray(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_DarkGreen(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_DarkMagenta(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.DarkMagenta;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_DarkRed(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.DarkRed;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_DarkYellow(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_Gray(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_Green(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_Magenta(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_White(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.White;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }
        public static void ConsOut_Yellow(XmlDocument xml)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            xml.Save(Console.Out);
            Console.WriteLine(Environment.NewLine);
            Console.ResetColor();
        }


        /// Error Logging
        public static void errorLog(string errNum, string errLoc, string errFull)
        {
            string q_error = @"INSERT INTO CT_UPS_QVErrorLog (errNum, errLoc, errFull)
                               VALUES ('" + errNum + "','" + errLoc + "','" + Globals.StringTool.RemoveApostrophes(errFull) + "')";
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.printableConnString))
                {
                    SqlCommand command = new SqlCommand(q_error, conn);
                    try
                    {
                        int rowsInserted = 0;
                        command.Connection.Open();
                        rowsInserted = command.ExecuteNonQuery();
                        command.Dispose();
                        command = null;
                        Console.WriteLine(rowsInserted.ToString() + " error logged - Error #" + errNum);
                    }
                    catch (Exception e)
                    {
                        errorLog("ERR-0", Globals.StringTool.Truncate(e.ToString(), 3999), "errorLog");
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("ERR-1", Globals.StringTool.Truncate(e.ToString(), 3999), "errorLog");
                Globals.Write_Magenta(e.ToString()); Console.Beep();
            }

            //Helpers help = new Helpers();
            //help.sendEmail(0, pErrNum + "<br/><br/>" + pErrMsg);
        }

        public static void errorLog(string errNum, string errLoc, string errFull, string refNum1, string refNum2)
        {
            string q_error = @"INSERT INTO CT_UPS_QVErrorLog (errNum, errLoc, errFull, refNum1, refNum2)
                               VALUES ('" + errNum + "','" + errLoc + "','" + Globals.StringTool.RemoveApostrophes(errFull) + "','" + refNum1 + "','" + refNum2 + "')";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.printableConnString))
                {
                    SqlCommand command = new SqlCommand(q_error, conn);
                    try
                    {
                        int rowsInserted = 0;
                        command.Connection.Open();
                        rowsInserted = command.ExecuteNonQuery();
                        command.Dispose();
                        command = null;
                        Console.WriteLine(rowsInserted.ToString() + " error logged - Error #" + errNum);
                    }
                    catch (Exception e)
                    {
                        errorLog("ERR-0", e.ToString(), "errorLog");
                        Globals.Write_Magenta(e.ToString()); Console.Beep();
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("ERR-1", e.ToString(), "errorLog");
                Globals.Write_Magenta(e.ToString()); Console.Beep();
            }

            //Helpers help = new Helpers();
            //help.sendEmail(0, pErrNum + "<br/><br/>" + pErrMsg);
        }
    }
}
