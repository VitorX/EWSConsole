using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace EWSConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //ImpersonationSample();
            SharedMailboxSample();
            Console.ReadLine();
        }

        static void ImpersonationSample()
        {
            string userName = "";
            string password = "";

            string impersonationUserName = "";
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new NetworkCredential(userName, password);

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;

            service.AutodiscoverUrl(userName, RedirectionUrlValidationCallback);

            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, impersonationUserName);

            Folder newFolder2 = new Folder(service);
            newFolder2.DisplayName = "TestFolder1";

            newFolder2.Save(WellKnownFolderName.Inbox);

        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        static void SharedMailboxSample()
        {
            string userName = "";
            string password = "";

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new NetworkCredential(userName, password);

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;

            service.AutodiscoverUrl(userName, RedirectionUrlValidationCallback);


            FolderId SharedMailbox = new FolderId(WellKnownFolderName.Inbox, "sharedmailboxFei@O365E3W15.onmicrosoft.com");
            ItemView itemView = new ItemView(10);
            var results = service.FindItems(SharedMailbox, itemView);
            foreach (var item in results)
            {
                Console.WriteLine(item.Subject);
            }

        }

    }
}
