using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace Mails_Console_Monthly
{
    class Program
    {
        static void Main(string[] args)
        {
            MailProcessor processor = new MailProcessor();
            processor.ProcessEmails();             
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
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    cmd.CommandText = "truncate table dbo.tbl_outlook_mails_monthly_dotnet_kycchecks_mumbai_mailbox_archive";
                    cmd.ExecuteNonQuery();
                }

            }
            catch (System.Exception ex)
            {

            }

            DateTime today = DateTime.Today;
            DateTime thirtyDaysAgo = today.AddDays(-30);
            string startDate = thirtyDaysAgo.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string endDate = today.ToString("dd/MM/yyyy 23:59", CultureInfo.InvariantCulture);

            Items filteredItems = null;
            Application outlookApp = null;
            NameSpace outlookNamespace = null;
            MAPIFolder mailbox = null;
            MAPIFolder inbox = null;

            try
            {
                outlookApp = new Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
                mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                inbox = mailbox.Folders["Inbox"];

                string filter = $"[ReceivedTime] >= '{startDate}' AND [ReceivedTime] <= '{endDate}'";
                filteredItems = inbox.Items.Restrict(filter);
                filteredItems.Sort("[ReceivedTime]", false);

                for (int i = 1; i <= filteredItems.Count; i++)
                {
                    MailItem mail = null;
                    AddressEntry senderEntry = null;
                    ExchangeUser exchUser = null;
                    Recipients recipients = null;

                    try
                    {
                        mail = filteredItems[i] as MailItem;
                        if (mail != null)
                        {
                            DateTime receivedtime = mail.ReceivedTime;
                            string subject = mail.Subject;
                            string sender = mail.SenderEmailAddress;
                            string cc = mail.CC;
                            string to = mail.To;
                            string categories = mail.Categories;
                            var importance = mail.Importance;
                            string entryid = mail.EntryID;
                            bool isunread = mail.UnRead;

                            bool isMarkedAsTask = mail.IsMarkedAsTask;
                            //string taskStatus = ((OlTaskStatus)mail.TaskStatus).ToString();
                            string flagStatus = string.Empty;
                            try
                            {
                                flagStatus = Enum.GetName(typeof(Microsoft.Office.Interop.Outlook.OlFlagStatus), mail.FlagStatus);
                            }
                            catch (System.Exception ex)
                            {
                                flagStatus = "Unknown";
                            }

                            using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                            {
                                conn.Open();
                                cmd.Parameters.Clear();
                                cmd.Connection = conn;
                                cmd.CommandText = "INSERT INTO dbo.tbl_outlook_mails_monthly_dotnet_kycchecks_mumbai_mailbox_archive (Subject,ReceivedDateTime,Sender,Categories,CC,Importance,EntryID,UploadDateTime,IsUnread,[TO],IsFlagged,FlagStatus) VALUES (@Subject,@ReceivedDateTime,@Sender,@Categories,@CC,@Importance,@EntryID,@UploadDateTime,@IsUnread,@TO,@IsFlagged,@FlagStatus)";
                                cmd.Parameters.AddWithValue("@Subject", subject ?? "");
                                cmd.Parameters.AddWithValue("@ReceivedDateTime", receivedtime);
                                cmd.Parameters.AddWithValue("@Sender", sender ?? "");
                                cmd.Parameters.AddWithValue("@Categories", categories ?? "");
                                cmd.Parameters.AddWithValue("@CC", cc ?? "");
                                cmd.Parameters.AddWithValue("@TO", to ?? "");
                                cmd.Parameters.AddWithValue("@Importance", importance);
                                cmd.Parameters.AddWithValue("@EntryID", entryid ?? "");
                                cmd.Parameters.AddWithValue("@UploadDateTime", DateTime.Now.ToLocalTime());
                                cmd.Parameters.AddWithValue("@IsUnread", isunread);
                                cmd.Parameters.AddWithValue("@IsFlagged", isMarkedAsTask);
                                cmd.Parameters.AddWithValue("@FlagStatus", flagStatus);
                                cmd.ExecuteNonQuery();
                            }

                        }
                    }
                    catch (System.Exception ex)
                    {

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
            catch (System.Exception)
            {

            }
            finally
            {
                // Release all top-level COM objects
                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                if (inbox != null) Marshal.ReleaseComObject(inbox);
                if (mailbox != null) Marshal.ReleaseComObject(mailbox);
                if (outlookNamespace != null) Marshal.ReleaseComObject(outlookNamespace);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);

                // Force garbage collection to clean up
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
                for (int i = 0; i < 3; i++)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    //cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "exec dbo.usp_reading_mails_outlook_dotnet_kycchecks_mumbai_mailbox_monthly";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (System.Exception ex)
            {

            }
        }
    }
}
