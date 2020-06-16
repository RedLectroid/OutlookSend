using System;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSend
{

    class Program
    {
        static int Main(string[] args)
        {

            int i = 0;
            string Subject = string.Empty;
            string Attachment = string.Empty;
            string Body = string.Empty;
            string Recipients = string.Empty;

            try
            {
                foreach (string argument in args)
                {
                    if (args[i].ToLower() == "-s" || args[i].ToLower() == "--subject")
                    {
                        Subject = args[i + 1];
                        Console.WriteLine("Subject is " + Subject);
                    }

                    if (args[i].ToLower() == "-a" || args[i].ToLower() == "--attachment")
                    {
                        Attachment = args[i + 1];
                        Console.WriteLine("Attachment is " + Attachment);
                    }

                    if (args[i].ToLower() == "-r" || args[i].ToLower() == "--recipients")
                    {
                        Recipients = args[i + 1];
                        Console.WriteLine("Recipients is " + Recipients);

                    }

                    if (args[i].ToLower() == "-b" || args[i].ToLower() == "--body")
                    {
                        Body = args[i + 1];
                        Console.WriteLine("Body is " + Body);

                    }
                    i++;
                };


                Boolean status = checkOutlook();
                foreach (string recipient in Recipients.Split(',').ToList())
                {
                    SendEmail(Subject, Attachment, recipient, Body);
                }
                Thread.Sleep(5000);
                if (status == true)
                {
                    killOutlook();
                }

                return 0;

            }

            catch (Exception ex)
            {
                return 1;
            }
        }

        static Boolean checkOutlook()
        {
            string OutlookPath = @"C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe";

            if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe"))
            {
                OutlookPath = @"C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe";
            }

            else if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office16\Outlook.exe"))
            {
                OutlookPath = @"C:\Program Files (x86)\Microsoft Office\Office16\Outlook.exe";
            }

            else if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\root\Office16\Outlook.exe"))
            {
                OutlookPath = @"C:\Program Files (x86)\Microsoft Office\root\Office16\Outlook.exe";
            }

            else
            {
                Console.WriteLine("Outlook not installed, exiting.");
                Environment.Exit(0);
            }

            Process[] pname = Process.GetProcessesByName("Outlook");
            Outlook.Application application = null;

            if (pname.Length == 0)
            {

                Console.WriteLine("Outlook not running, I will start it for you minimized\n");
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("Outlook", "", Missing.Value, Missing.Value);
                nameSpace = null;
                Thread.Sleep(10000);
                return true;
            }
            else
            {
                return false;
            }
        }

        static void killOutlook()
        {
            foreach (var process in Process.GetProcessesByName("Outlook"))
            {
                process.Kill();
                Console.WriteLine("Killing minimized Outlook process");
            }
        }

        static void SendEmail(string subject, string attachment, string recipient, string body)
        {
            var emailTemplate = OutlookSend.Properties.Resource1.emailBody;
            Outlook.Application Application = new Outlook.Application();
            Outlook.MailItem mail = Application.CreateItem(
                Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = subject;
            if (body == "HTML" || body == "html")
            {
                mail.HTMLBody = emailTemplate;
            }
            else
            {
                mail.Body = body;
            }

            Outlook.AddressEntry currentUser =
                Application.Session.CurrentUser.AddressEntry;
            if (currentUser.Type == "EX")
            {

                Console.WriteLine("Sending email to " + recipient);
                mail.Recipients.Add(recipient);
                mail.Recipients.ResolveAll();
                if (attachment != "")
                {
                    mail.Attachments.Add(attachment,
                    Outlook.OlAttachmentType.olByValue, Type.Missing,
                        Type.Missing);
                }
                mail.Send();
                Thread.Sleep(2000);


            }
        }

    }
}