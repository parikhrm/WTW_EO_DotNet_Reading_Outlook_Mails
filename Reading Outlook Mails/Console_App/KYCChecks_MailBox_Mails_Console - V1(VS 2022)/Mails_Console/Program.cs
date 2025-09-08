using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
//using System.Windows.Forms;
using System.Data.SqlClient;
//new ones
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
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
            MailProcessor processor = new MailProcessor();
            processor.ProcessEmails();
            //Console.ReadKey();
        }
    }

    public class MailProcessor
    {
        private string connectionstringtxt = "Data Source=A20-CB-DBSE01P;Initial Catalog=DRD;User ID=DRDUsers;Password=24252425";

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
                    //Console.WriteLine("Daily table truncated successfully.");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error truncating table: {ex.Message}");
            }


            DateTime today = DateTime.Today;
            DateTime thirtyDaysAgo = today.AddDays(-2);
            string startDate = thirtyDaysAgo.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string endDate = today.ToString("dd/MM/yyyy 23:59", CultureInfo.InvariantCulture);

            Items filteredItems = null;
            Application outlookApp = null;
            NameSpace outlookNamespace = null;
            MAPIFolder mailbox = null;
            MAPIFolder inbox = null;
            Items items = null;

            try
            {
                outlookApp = new Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
                mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                inbox = mailbox.Folders["Inbox"];

                string filter = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'";
                filteredItems = inbox.Items.Restrict(filter);
                filteredItems.Sort("[ReceivedTime]", true);

                Console.WriteLine($"Processing {filteredItems.Count} emails...");

                foreach (object item in filteredItems)
                {
                    MailItem mail = item as MailItem;
                    if (mail != null)
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

                                    }
                                    if (senderEntry != null) Marshal.ReleaseComObject(senderEntry);
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
                            string toRecipients = GetRecipientEmailAddresses(mail.Recipients);

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
                //MessageBox.Show("Error Generated Details: " + ab.ToString());
            }
        }

        private string GetRecipientEmailAddresses(Recipients recipients)
        {
            var recipientList = new List<string>();

            if (recipients == null) return "";

            foreach (Recipient recipient in recipients)
            {
                try
                {
                    recipient.Resolve();
                    if (recipient.Resolved)
                    {
                        AddressEntry addressEntry = recipient.AddressEntry;
                        if (addressEntry != null)
                        {
                            // For Exchange users, get the primary SMTP address
                            if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
                                addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                            {
                                if (addressEntry.GetExchangeUser() != null)
                                {
                                    recipientList.Add(addressEntry.GetExchangeUser().PrimarySmtpAddress);
                                }
                                else if (addressEntry.GetExchangeDistributionList() != null)
                                {
                                    recipientList.Add(addressEntry.GetExchangeDistributionList().PrimarySmtpAddress);
                                }
                            }
                            else
                            {
                                // For other address types, use the general address property
                                recipientList.Add(addressEntry.Address);
                            }
                            Marshal.ReleaseComObject(addressEntry);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine($"Error resolving recipient: {ex.Message}");
                }
                finally
                {
                    if (recipient != null) Marshal.ReleaseComObject(recipient);
                }
            }

            return string.Join(";", recipientList);
        }
    }
}
