using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadOffice365Mailbox
{
    class Program
    {
        private static String recipients;

        static void Main(string[] args)
        {
            ExchangeService _service;

            try
            {
                Console.WriteLine("Exchange connection");

                _service = new ExchangeService
                {
                    Credentials = new WebCredentials("Vidhya.N@philips.com", "Lovepanda@97")
                };
            }
            catch
            {
                Console.WriteLine("new ExchangeService failed. Press enter to exit:");
                return;
            }

            // This is the office365 webservice URL
            _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            FolderId inboxId = new FolderId(WellKnownFolderName.Inbox, "Vidhya.N@philips.com");
            var findResults = _service.FindItems(inboxId, new ItemView(50));
            
            try
            {
                foreach (var message in findResults.Items)
                {
                    var msg = EmailMessage.Bind(_service, message.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));
                    foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in msg.Attachments)
                    {
                        Console.WriteLine("Reading mail");
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        // Load the file attachment into memory and print out its file name.
                        fileAttachment.Load();
                        var filename = fileAttachment.Name;
                        // Read 50 mails
                        foreach (EmailMessage email in _service.FindItems(WellKnownFolderName.Inbox, new ItemView(50)))
                        {
                            email.Load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));

                            // Then you can retrieve extra information like this:
                            recipients = "";
                           
                            String s1 = "RE: Training Sessions";
                            String s2 = email.Subject;
                            
                            if (s2.Equals(s1))
                            {
                                if (email.HasAttachments)
                                {
                                    Console.WriteLine("@message_id :" + email.InternetMessageId);
                                    Console.WriteLine("@from: " + email.From.Address.ToString());
                                    Console.WriteLine("@body:" + email.TextBody);
                                    Console.WriteLine("@SUbject:" + email.Subject);
                                    Console.WriteLine("Download Attachment- Success");
                                    var thestream = new FileStream("C:\\data\\download\\" + fileAttachment.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                                    fileAttachment.Load(thestream);
                                    thestream.Close();
                                    thestream.Dispose();
                                }
                            }

                        }

                        Console.WriteLine("Exiting");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("An error has occured. \n:" + e.Message);
            }
            Console.ReadLine();
        }
     
        }
    }

 
