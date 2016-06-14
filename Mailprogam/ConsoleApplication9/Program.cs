using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;
using System.IO.IsolatedStorage;
using System.Threading.Tasks;
using System.Net.Mail; 
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;


namespace ConsoleApplication9
{
    class Program
    {
        static string p4pMail = ConfigurationManager.AppSettings.Get("username");
        static string p4pPassword = ConfigurationManager.AppSettings.Get("password");
        
        static void Main(string[] args)
        {
            // maakt een nieuw instantie aan van de mail class
            mail mailClient = new mail(p4pMail);
            ImapClient client = mailClient.connect(p4pMail, p4pPassword);
            // selecteert de Inbox 
            IMailFolder inbox = client.Inbox;

            // loop door alle berichten in Inbox
            for (int x = 0; x < inbox.Count; x++)
            {
                var message = inbox.GetMessage(x);
                // als het onderwerp niet doldoet
                if(!mailClient.subjectCheck(message))
                {
                    // verwijder het bericht uit de inbox 
                    mailClient.deleteMessage(inbox, x); 
                }
                // als het onderwerp wel voldoet 
                // als het bericht 1 of meerdere bijlages heeft
                else if(message.Attachments.Count() > 0) 
                {
                    // sla de bijlage op 
                    mailClient.saveAttachment(message);
                    // stuurt een ontvangst mail
                    mailClient.sendMessage(message, ConfigurationManager.AppSettings.Get("compleetMessage")); 
                // als het bericht geen bijlage heeft
                } else 
                {
                    // stuurt een niet compleet mail 
                    mailClient.sendMessage(message, ConfigurationManager.AppSettings.Get("notCompleetMessage")); 
                }
                // verwijder de mail uit de inbox
                mailClient.deleteMessage(inbox, x);
            }
            // disconnect de mailclients 
             mailClient.disconnect();
        }
    }
}
