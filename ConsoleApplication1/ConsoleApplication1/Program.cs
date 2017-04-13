using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ImapX;
using ImapX.Enums;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Data.OleDb;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Weather();
        }

        private static string CheckSubject(string Subject)
        {
            Subject = Subject.ToLower();
            Subject = Regex.Replace(Subject, "[-.?!)(,: }{1234567890]", "");
            if ((Subject == "weather") || (Subject == "forecast") || (Subject == "forecast/daily"))
            {

            }
            else
            {
                Subject = "false";
            }
            return Subject;
        }

        public static void SendMail(string smtpServer, string from, string password, string mailto, string caption, string message, string attachFile = null)
        {
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new System.Net.Mail.MailAddress(from);
                mail.To.Add(new System.Net.Mail.MailAddress(mailto));
                mail.Subject = caption;
                mail.Body = message;
                if (!string.IsNullOrEmpty(attachFile))
                    mail.Attachments.Add(new System.Net.Mail.Attachment(attachFile));
                SmtpClient client = new SmtpClient();
                client.Host = smtpServer;
                client.Port = 587;
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(from.Split('@')[0], password);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(mail);
                mail.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Mail.Send: " + e.Message);
            }
        }

        private static void Weather()
        {
            string City = "";
            string[] Email = File.ReadAllLines(@"Email.txt");
            WeatherTxt weatherTxt = new WeatherTxt();

            ImapClient client = new ImapClient("imap.gmail.com", true);

            if (client.Connect())
            {
                if (client.Login("nublikaskaweather@gmail.com", "123456789dD"))
                {
                    client.Folders.Inbox.Messages.Download("ALL", MessageFetchMode.Full, 2);
                    foreach (Message message in client.Folders.Inbox.Messages)
                    {
                        Console.WriteLine(ReadFromExcel(message.From.Address.ToString()) );
                        if ((message.Seen == false) && (CheckSubject(message.Subject) != "false") && (ReadFromExcel(message.From.Address.ToString())))
                        {
                            message.Seen = true;
                            Console.WriteLine(message.From.Address.ToString());
                            City = Regex.Replace(message.Body.Text, "[-.?!)(,: }{1234567890]", "");
                            weatherTxt.GetWeather(CheckSubject(message.Subject), City);
                            SendMail("smtp.gmail.com",
                                     Email[0],
                                     Email[1],
                                     message.From.Address.ToString(),
                                     CheckSubject(message.Subject),
                                     weatherTxt.getBody(),
                                     weatherTxt.GetDate());
                            Console.WriteLine("сообщение отправлено");
                            WriteToExcel(message);
                            weatherTxt.Clear();
                        }
                    }
                }
                int a = Console.Read();
            }
        }

        private static bool ReadFromExcel(string address)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(@"H:\Users\denis\Documents\Visual Studio 2015\Projects\ConsoleApplication1\ConsoleApplication1\bin\Debug", "emails.xlsx") + ";" +
                                    "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=0;MAXSCANROWS=0'";
            conn.Open();
            string sheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null)?.Rows[1]["TABLE_NAME"].ToString();
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
            cmd.Connection = conn;
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                Console.WriteLine(reader[0].ToString());
                if (reader[0].ToString() == address)
                {
                    reader.Close();
                    conn.Close();
                    return true;
                }
            }
            return false;
        }

        private static void WriteToExcel(ImapX.Message message)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(@"H:\Users\denis\Documents\Visual Studio 2015\Projects\ConsoleApplication1\ConsoleApplication1\bin\Debug", "emails.xlsx") + ";" +
                                    "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=0;MAXSCANROWS=0'";
            conn.Open();
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            string sheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null)?.Rows[0]["TABLE_NAME"].ToString();
            cmd.CommandText = $"INSERT INTO [{sheet}] ([Date], [Email], [Subject], [City]) VALUES('{message.Date}', '{message.From.Address}', '{message.Subject}', '{message.Body.Text}')";
            cmd.ExecuteNonQuery();
            conn.Close();
        }
    }

}
