using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;

namespace ExchangeBulkMailAttach
{
    class Program
    {
        public static string mailHost = ConfigurationManager.AppSettings["mailhost"];
        public static string mailUser = ConfigurationManager.AppSettings["mailuser"];
        public static string mailPass = ConfigurationManager.AppSettings["mailpass"];
        public static string documentPath = ConfigurationManager.AppSettings["documentpath"];
        static void Main(string[] args)
        {
            /*proxy behind*/
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
                  service.UseDefaultCredentials = true;
                  service.TraceEnabled = true;
                  service.TraceFlags = TraceFlags.All;
                  service.WebProxy = WebRequest.GetSystemWebProxy();
                  service.Url = new Uri(mailHost);
                  EmailMessage email = new EmailMessage(service);
                  service.Credentials = new WebCredentials(mailUser, mailPass);
                  IWebProxy proxy = WebRequest.GetSystemWebProxy();
                  proxy.Credentials = CredentialCache.DefaultCredentials;
                  service.WebProxy = proxy;

            /*auto discovery 
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Credentials = new WebCredentials("i.tunga@hotmail.com", "pass");
            service.AutodiscoverUrl("i.tunga@hotmail.com");
            */
            string[] pdfFiles = Directory.GetFiles(documentPath, "*.txt")
                                     .Select(Path.GetFileNameWithoutExtension)
                                     .ToArray();
            
            foreach (string userName in pdfFiles)
            {
                string mailAddress = userName + "@gmail.com";
                EmailMessage message = new EmailMessage(service);
                message.Subject = "Interesting";
                message.Body = "The proposition has been considered.";
                message.ToRecipients.Add(mailAddress);
                message.Attachments.AddFileAttachment(documentPath + userName + ".txt");
                message.SendAndSaveCopy();
            }
        }
    }
}
