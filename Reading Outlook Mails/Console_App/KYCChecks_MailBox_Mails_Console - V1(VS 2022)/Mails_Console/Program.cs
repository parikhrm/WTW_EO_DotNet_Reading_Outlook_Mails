using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
//using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;
//new ones
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
//using System.Windows.Forms.Integration;
//using System.Windows.Forms.Design;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Outlook = Microsoft.Office.Interop.Outlook;




namespace Mails_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            //new comments
            MailProcessor processor = new MailProcessor();
            processor.ProcessEmails();
            //Console.ReadKey();
        }
    }
    /*
    public static class OutlookHelper
    {
        // Imports for Win32 API functions to send messages to windows
        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        private const int WM_QUIT = 0x12;

        public static void KillOutlook()
        {
            // Method 1: Ask Outlook nicely to close
            Process[] processes = Process.GetProcessesByName("OUTLOOK");
            foreach (Process process in processes)
            {
                try
                {
                    // Send a WM_QUIT message to the main window
                    SendMessage(process.MainWindowHandle, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                    process.WaitForExit(5000); // Wait for up to 5 seconds
                    if (!process.HasExited)
                    {
                        process.Kill();
                    }
                }
                catch (System.Exception ex)
                {
                    //Console.WriteLine($"Error closing Outlook process: {ex.Message}");
                }
            }

            // Method 2: Brute force kill any remaining processes
            processes = Process.GetProcessesByName("OUTLOOK");
            foreach (Process process in processes)
            {
                try
                {
                    process.Kill();
                }
                catch (System.Exception ex)
                {
                    //Console.WriteLine($"Error killing Outlook process: {ex.Message}");
                }
            }
        }
    }
    */
    public class MailProcessor
    {
        private string connectionstringtxt = "Data Source=A20-CB-DBSE01P;Initial Catalog=DRD;User ID=DRDUsers;Password=24252425";

        // Helper method to properly release COM objects
        private void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
            }
        }
        private string GetRecipientAddressesFromCollection(Recipients recipients, OlMailRecipientType recipientType)
        {
            if (recipients == null || recipients.Count == 0)
            {
                return string.Empty;
            }

            var addressList = new List<string>();
            // COM collections are 1-based, so iterate from 1
            for (int i = 1; i <= recipients.Count; i++)
            {
                Recipient recipient = recipients[i];
                try
                {
                    if (recipient.Type == (int)recipientType)
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        string smtpAddress = string.Empty;

                        // Use PropertyAccessor for reliable SMTP address lookup
                        try
                        {
                            smtpAddress = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS).ToString();
                        }
                        catch (System.Exception ex)
                        {
                            // Fallback to the regular Address property, which can be an internal address
                            smtpAddress = recipient.Address;
                            Console.WriteLine($"An error occurred during Outlook processing: {ex.Message}");
                        }
                        if (!string.IsNullOrEmpty(smtpAddress))
                        {
                            addressList.Add(smtpAddress);
                        }
                    }
                }
                finally
                {
                    //if (recipient != null) Marshal.ReleaseComObject(recipient);
                    ReleaseComObject(recipient);
                }
            }

            if (recipients != null) Marshal.ReleaseComObject(recipients);

            return string.Join(", ", addressList);
        }
        public void ProcessEmails()
        {
            SqlCommand cmd = new SqlCommand();

            try
            {
                // Truncate daily table
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    cmd.CommandText = "truncate table dbo.tbl_outlook_mails_daily_dotnet_kycchecks_mumbai_mailbox";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (System.Exception ex)
            {
                //onsole.WriteLine($"Error truncating table: {ex.Message}");
                Console.WriteLine($"An error occurred during Outlook processing: {ex.Message}");
            }


            DateTime today = DateTime.Today;
            DateTime fiveDaysAgo = today.AddDays(-5);
            string startDate = fiveDaysAgo.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string endDate = today.ToString("dd/MM/yyyy 23:59", CultureInfo.InvariantCulture);
            //string startDate = fiveDaysAgo.ToString("yyyy-MM-dd HH:mm");
            //string endDate = today.ToString("yyyy-MM-dd HH:mm");

            Items filteredItems = null;
            Application outlookApp = null;
            NameSpace outlookNamespace = null;
            MAPIFolder mailbox = null;
            MAPIFolder inbox = null;

            //new code
            try
            {
                outlookApp = new Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
                mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                inbox = mailbox.Folders["Inbox"];

                string filter = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'";
                filteredItems = inbox.Items.Restrict(filter);
                filteredItems.Sort("[ReceivedTime]", false);


                // Use a 'for' loop to iterate and release each mail item explicitly.
                for (int i = 1; i <= filteredItems.Count; i++)
                {
                    MailItem mail = null;
                    AddressEntry senderEntry = null;
                    ExchangeUser exchUser = null;
                    Recipients recipients = null;

                    try
                    {
                        mail = filteredItems[i] as MailItem;

                        if (mail != null && mail.FlagStatus == Microsoft.Office.Interop.Outlook.OlFlagStatus.olNoFlag)
                        {
                            DateTime receivedtime = mail.ReceivedTime;
                            string subject = mail.Subject;
                            //string cc = mail.CC;
                            string categories = mail.Categories;
                            var importance = mail.Importance;
                            string entryid = mail.EntryID;
                            bool isunread = mail.UnRead;
                            bool isMarkedAsTask = mail.IsMarkedAsTask;
                            string flagStatus = Enum.GetName(typeof(OlFlagStatus), mail.FlagStatus);

                            //string flagStatus = "Unknown";
                            try
                            {
                                flagStatus = Enum.GetName(typeof(Microsoft.Office.Interop.Outlook.OlFlagStatus), mail.FlagStatus);
                            }
                            catch (System.Exception ex)
                            {
                                Console.WriteLine($"An error occurred during Outlook processing: {ex.Message}");
                            }

                            string senderEmail = string.Empty;
                            senderEntry = mail.Sender;
                            //try
                            //{
                                //if (mail.SenderEmailType == "EX")
                                //{
                                    //senderEntry = mail.Sender;
                                    //if (senderEntry != null)
                                    //{
                                    //    exchUser = senderEntry.GetExchangeUser();
                                    //    if (exchUser != null && !string.IsNullOrEmpty(exchUser.PrimarySmtpAddress))
                                    //    {
                                    //        senderEmail = exchUser.PrimarySmtpAddress;
                                    //    }
                                    //    else
                                    //    {
                                    //        senderEmail = senderEntry.Address;
                                    //    }
                                    //}
                                //}
                                //else
                                //{
                                    //senderEmail = mail.SenderEmailAddress;
                                //}
                            //}
                            //catch (System.Exception ex)
                            //{
                                //senderEmail = mail.SenderEmailAddress; // fallback
                                //Console.WriteLine($"An error occurred during Outlook processing: {ex.Message}");
                            //}

                            //recipients = mail.Recipients;
                            //string toRecipients = GetRecipientAddressesFromCollection(recipients, OlMailRecipientType.olTo);
                            //ReleaseComObject(recipients); // Release here immediately after use

                            using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                            {
                                conn.Open();
                                cmd.Parameters.Clear();
                                cmd.Connection = conn;
                                cmd.CommandText = "INSERT INTO dbo.tbl_outlook_mails_daily_dotnet_kycchecks_mumbai_mailbox " +
                                                  "(Subject,ReceivedDateTime,Sender,Categories,CC,Importance,EntryID,UploadDateTime,IsUnread,[To],IsFlagged,FlagStatus) " +
                                                  "VALUES (@Subject,@ReceivedDateTime,@Sender,@Categories,@CC,@Importance,@EntryID,@UploadDateTime,@IsUnread,@TO,@IsFlagged,@FlagStatus)";
                                cmd.Parameters.AddWithValue("Subject", (object)subject ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@ReceivedDateTime", receivedtime);
                                cmd.Parameters.AddWithValue("@Sender", (object)senderEmail ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@Categories", (object)categories ?? DBNull.Value);
                                //cmd.Parameters.AddWithValue("@CC", (object)cc ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@CC", DBNull.Value);
                                cmd.Parameters.AddWithValue("@Importance", importance);
                                cmd.Parameters.AddWithValue("@EntryID", (object)entryid ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@UploadDateTime", DateTime.Now.ToLocalTime());
                                cmd.Parameters.AddWithValue("@IsUnread", isunread);
                                //cmd.Parameters.AddWithValue("@TO", (object)toRecipients ?? DBNull.Value);
                                cmd.Parameters.AddWithValue("@TO", DBNull.Value);
                                cmd.Parameters.AddWithValue("@IsFlagged", isMarkedAsTask);
                                cmd.Parameters.AddWithValue("@FlagStatus", (object)flagStatus ?? DBNull.Value);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (System.Exception mailEx)
                    {
                        Console.WriteLine($"Error processing mail item: {mailEx.Message}");
                    }
                    finally
                    {
                        // Release COM objects in reverse order of their creation
                        if (recipients != null) Marshal.ReleaseComObject(recipients);
                        if (exchUser != null) Marshal.ReleaseComObject(exchUser);
                        if (senderEntry != null) Marshal.ReleaseComObject(senderEntry);
                        if (mail != null) Marshal.ReleaseComObject(mail);


                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"An error occurred during Outlook processing: {ex.Message}");
            }
            finally
            {
                // Release all top-level COM objects
                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                if (inbox != null) Marshal.ReleaseComObject(inbox);
                if (mailbox != null) Marshal.ReleaseComObject(mailbox);
                if (outlookNamespace != null) Marshal.ReleaseComObject(outlookNamespace);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);



                //Force garbage collection to clean up
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
                for (int i = 0; i < 3; i++)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

            //commenting for each loop(*)
            /*
            try
            {
                outlookApp = new Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
                mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                inbox = mailbox.Folders["Inbox"];

                // The DASL query for unflagged items uses the `FlagStatus` property, which has a
                // numerical value of 0 for `olNoFlag`.
                // The property is 'http://schemas.microsoft.com/mapi/proptag/0x10900003' or
                // a GUID-based one. Using the `FlagStatus` property name is a simpler approach
                // for most cases and maps to the correct MAPI property.

                string filter = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'";
                filteredItems = inbox.Items.Restrict(filter);
                filteredItems.Sort("[ReceivedTime]", false);

                //Console.WriteLine($"Processing {filteredItems.Count} emails...");

                foreach (object item in filteredItems)
                {
                    MailItem mail = item as MailItem;
                    if (mail != null)
                    {
                        // ADD THIS CONDITION: Check if the mail's flag status is not flagged
                        if (mail.FlagStatus == Microsoft.Office.Interop.Outlook.OlFlagStatus.olNoFlag)
                        {
                            try
                            {
                                DateTime receivedtime = mail.ReceivedTime;
                                string subject = mail.Subject;
                                string cc = mail.CC;
                                string categories = mail.Categories;
                                var importance = mail.Importance;
                                string entryid = mail.EntryID;
                                bool isunread = mail.UnRead;

                                bool isMarkedAsTask = mail.IsMarkedAsTask;
                                //string taskStatus = ((OlTaskStatus)mail.TaskStatus).ToString();
                                string flagStatus = string.Empty;
                                try
                                {
                                    // The correct property for MailItem flag status is FlagStatus
                                    Microsoft.Office.Interop.Outlook.OlFlagStatus olFlagStatus = mail.FlagStatus;
                                    flagStatus = Enum.GetName(typeof(Microsoft.Office.Interop.Outlook.OlFlagStatus), olFlagStatus);
                                }
                                catch
                                {
                                    // Handle cases where the status is not a standard value
                                    flagStatus = "Unknown";
                                }

                                string senderEmail = string.Empty;
                                try
                                {
                                    if (mail.SenderEmailType == "EX")
                                    {
                                        Outlook.AddressEntry senderEntry = mail.Sender;
                                        if (senderEntry != null)
                                        {
                                            Outlook.ExchangeUser exchUser = senderEntry.GetExchangeUser();
                                            if (exchUser != null && !string.IsNullOrEmpty(exchUser.PrimarySmtpAddress))
                                            {
                                                senderEmail = exchUser.PrimarySmtpAddress;
                                            }
                                            else
                                            {
                                                senderEmail = senderEntry.Address;
                                            }
                                            if (exchUser != null) Marshal.ReleaseComObject(exchUser);
                                            //Marshal.ReleaseComObject(exchUser);

                                        }
                                        if (senderEntry != null) Marshal.ReleaseComObject(senderEntry);
                                        //Marshal.ReleaseComObject(senderEntry);
                                    }
                                    else
                                    {
                                        senderEmail = mail.SenderEmailAddress;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    // Optional: log or handle the exception
                                    senderEmail = mail.SenderEmailAddress; // fallback
                                }

                                // Call the helper method to get recipient email addresses
                                //string toRecipients = GetRecipientEmailAddresses(mail.Recipients);

                                // FIX: Get only 'To' addresses using the new helper method
                                string toRecipients = GetRecipientAddressesFromCollection(mail.Recipients, OlMailRecipientType.olTo);



                                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                                {
                                    conn.Open();
                                    cmd.Parameters.Clear();
                                    cmd.Connection = conn;
                                    cmd.CommandText = "INSERT INTO dbo.tbl_outlook_mails_daily_dotnet_kycchecks_mumbai_mailbox " +
                                                      "(Subject,ReceivedDateTime,Sender,Categories,CC,Importance,EntryID,UploadDateTime,IsUnread,[To],IsFlagged,FlagStatus) " +
                                                      "VALUES (@Subject,@ReceivedDateTime,@Sender,@Categories,@CC,@Importance,@EntryID,@UploadDateTime,@IsUnread,@TO,@IsFlagged,@FlagStatus)";
                                    cmd.Parameters.AddWithValue("Subject", subject ?? "");
                                    cmd.Parameters.AddWithValue("@ReceivedDateTime", receivedtime);
                                    cmd.Parameters.AddWithValue("@Sender", senderEmail ?? "");
                                    cmd.Parameters.AddWithValue("@Categories", categories ?? "");
                                    cmd.Parameters.AddWithValue("@CC", cc ?? "");
                                    cmd.Parameters.AddWithValue("@Importance", importance);
                                    cmd.Parameters.AddWithValue("@EntryID", entryid ?? "");
                                    cmd.Parameters.AddWithValue("@UploadDateTime", DateTime.Now.ToLocalTime());
                                    cmd.Parameters.AddWithValue("@IsUnread", isunread);
                                    cmd.Parameters.AddWithValue("@TO", toRecipients ?? "");
                                    cmd.Parameters.AddWithValue("@IsFlagged", isMarkedAsTask);
                                    cmd.Parameters.AddWithValue("@FlagStatus", flagStatus);
                                    cmd.ExecuteNonQuery();
                                }

                            }
                            catch (System.Exception mailEx)
                            {
                                Console.WriteLine($"Error processing mail item: {mailEx.Message}");
                            }
                            finally
                            {
                                if (mail != null) Marshal.ReleaseComObject(mail);
                                //Marshal.ReleaseComObject(mail);
                            }
                        } // End of the if (mail.FlagStatus...) block
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"An error occurred during Outlook processing: {ex.Message}");
            }
            finally
            {
                // Release COM objects to prevent memory leaks
                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                if (items != null) Marshal.ReleaseComObject(items);
                if (inbox != null) Marshal.ReleaseComObject(inbox);
                if (mailbox != null) Marshal.ReleaseComObject(mailbox);
                if (outlookNamespace != null) Marshal.ReleaseComObject(outlookNamespace);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
            */
            //commenting ends here (*)

            try
            {
                //run stored procedure
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    //cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "exec dbo.usp_reading_mails_outlook_dotnet_kycchecks_mumbai_mailbox";
                    //MessageBox.Show("Emails successfully saved to database");
                    cmd.ExecuteNonQuery();
                }
            }
            catch (SystemException ab)
            {
                //Console.WriteLine("Error Generated Details: " + ab.ToString());
            }
        }
    }
}
