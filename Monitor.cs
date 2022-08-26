using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace EmailCheckerService
{
    public class Monitor
    {
        public Monitor()
        {

        }

        public void Run()
        {
            ReadMailItems();
        }

        DateTime date = DateTime.Today;

        private void ReadMailItems()
        {
            Application outlookApplication = new Application();
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;

            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items; 

                int i = mailItems.Count;
                int x = 0;

                while (i>0)
                {
                    if(x<20)
                    {
                        if (mailItems[i].SenderEmailAddress == Constants.SenderAddress || mailItems[i].Subject == Constants.EmailSubject)
                        {
                            DateTime dateTimeOfEmail = mailItems[i].ReceivedTime;
                            string dateOfEmail = dateTimeOfEmail.ToShortDateString();
                            if (DateTime.Today.ToShortDateString() == dateOfEmail)
                            {
                                forwardEmail(mailItems[i]);
                                Marshal.ReleaseComObject(mailItems[i]);
                                return;
                            }
                        }
                        Marshal.ReleaseComObject(mailItems[i]);
                        i--;
                    }
                    x++;
                }
            }
            catch {
            
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        private static void forwardEmail(MailItem item)
        {
            item.Forward();
            item.Recipients.Add("enarod@gmail.com");
            item.Send();
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
