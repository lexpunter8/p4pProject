using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using System.IO;

namespace ConsoleApplication9
{
    class mail
    {
        private ImapClient receiveClient;
        private SmtpClient mailClient;
        private string p4pMail;
        /// <summary>
        /// constructor die een nieuwe smtp client en imapclient aan maakt
        /// </summary>
        /// <param name="p4pMail"></param>
        public mail (string p4pMail)
        {
            this.p4pMail = p4pMail;
            receiveClient = new ImapClient();
            mailClient = new SmtpClient();
        }

        /// <summary>
        /// schrijft alle applicatie errors naar een txt bestand en eindig het procces
        /// wanner er zich een error voor doet 
        /// </summary>
        /// <param name="text">exeption</param>
        private void writeLog(string text)
        {
            // als het bestand logfile.txt nog niet bestaat
            if(!File.Exists("logFile.txt"))
            {
                // maak het logfile.txt bestand aan 
                 var logf = File.CreateText("logFile.txt");
                 logf.Close();
            }

            using (StreamWriter tw = new StreamWriter("logFile.txt", true))
            {
                // schrijf in het text log bestand 
                tw.WriteLine("-------------------------------------------------------");
                tw.WriteLine(DateTime.Now);
                tw.WriteLine(text);
                // open text bestand met de errors 
                System.Diagnostics.Process.Start("logFile.txt");
            }
            // sluit applicatie
            Environment.Exit(Environment.ExitCode);
        }

        /// <summary>
        /// maakt een connectie met de imap en smtp clients
        /// </summary>
        /// <param name="mail"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public ImapClient connect(string mail, string password)
        {
            try
            {
                //imap client instelingen 
                string server = "Imap." + ConfigurationManager.AppSettings.Get("serverName") + ".com";
                receiveClient.Connect(server, 993, true);
                receiveClient.AuthenticationMechanisms.Remove("XOAUTH@");
                receiveClient.Authenticate(mail, password);
                receiveClient.Inbox.Open(FolderAccess.ReadWrite);
                //smtp client instellingen
                server = "smtp." + ConfigurationManager.AppSettings.Get("serverName") + ".com";
                mailClient.EnableSsl = true;
                mailClient.Host = server;
                mailClient.Port = 587;
                mailClient.UseDefaultCredentials = false;
                mailClient.Credentials = new System.Net.NetworkCredential(p4pMail, password);
                mailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            } catch (Exception e)
            {
                writeLog(Convert.ToString(e));
            }
            return receiveClient;
        }

        /// <summary>
        /// controleerd of het onderwerp van het bericht voldoet 
        /// </summary>
        /// <param name="message"></param>
        /// <returns>true/false</returns>
        public bool subjectCheck(MimeMessage message, int count = 0)
        {
            if (count >= ConfigurationManager.AppSettings.Count) return false;
            string key = ConfigurationManager.AppSettings.GetKey(count);
            if(key.StartsWith("subjectSearch"))
            {
                string Apps = ConfigurationManager.AppSettings.Get(key);
                if(message.Subject.Length >= Apps.Length)
                {
                    if (message.Subject.Substring(0, Apps.Length).ToLower() == Apps.ToLower()) return true;
                    else return false || subjectCheck(message, count + 1);
                }
                else return false || subjectCheck(message, count + 1);
            }
            else return subjectCheck(message, count + 1);
        }

        /// <summary>
        /// slaat de bijlage van het bericht op
        /// </summary>
        /// <param name="message"></param>
        public void saveAttachment(MimeMessage message)
        {
            try
            {
                // ga naar de huidige map
                string path = System.IO.Directory.GetCurrentDirectory();
                // loop door de bijlages heen en sla die op
                foreach (MimePart attachment in message.Attachments)
                {
                    // pad waar de bijlage wordt opgeslagen
                    string pathToSave = ConfigurationManager.AppSettings.Get("pathtosave");
                    string file = path + pathToSave + attachment.FileName;// extentie
                    using (FileStream stream = File.Open(file, FileMode.Create))
                    {
                        attachment.ContentObject.DecodeTo(stream);
                    }
                }
            } catch(Exception e)
            {
                writeLog(Convert.ToString(e));
            }
        }

        /// <summary>
        /// verwijderd het bericht 
        /// </summary>
        /// <param name="inbox"></param>
        /// <param name="x">nummer van het bericht wat verwijderd moet worden</param>
        public void deleteMessage(IMailFolder inbox, int x)
        {
            inbox.AddFlags(x, MessageFlags.Deleted, false);
        }

        /// <summary>
        /// verstuurt een mail 
        /// </summary>
        /// <param name="message">het bericht waar op gereageerd moet worden</param>
        /// <param name="body">inhoud van het bericht</param>
        public void sendMessage(MimeMessage message, string body)
        {
            // vervang "$$" voor "<br/>" 
            body = body.Replace("$$", "<br/>");
            // voeg de datum van het ontvangen bericht toe 
            body = string.Format("Betreft het bericht ontvangen op: {0} <br/>", message.Date.ToLocalTime()) + body; 
            MailMessage m = new MailMessage();
            m.From = new MailAddress(p4pMail);
            m.To.Add(new MailAddress(Convert.ToString(message.From)));
            m.Subject = string.Format("RE: {0}", message.Subject);
            m.Body = body;
            m.IsBodyHtml = true;
            try
            {
                // verstuur de mail
                mailClient.Send(m);

            } catch (Exception e)
            {
                writeLog(Convert.ToString(e));
            }
        }

        public void disconnect()
        {
            receiveClient.Disconnect(true);
        }
    }
}
