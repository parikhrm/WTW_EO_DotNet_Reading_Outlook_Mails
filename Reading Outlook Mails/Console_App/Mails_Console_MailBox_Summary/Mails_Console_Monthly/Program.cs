using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//new ones
using System.Drawing;
//using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Configuration;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
//using System.Windows.Forms.Integration;
//using System.Windows.Forms.Design;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Timers;

namespace Mails_Console_Monthly
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionstringtxt = "Data Source=A20-CB-DBSE01P;Initial Catalog=DRD;User ID=DRDUsers;Password=24252425";
            //public string connectionstringtxt = ConfigurationManager.ConnectionStrings["KYC_RDC_Workflow.Properties.Settings.DRDConnectionString"].ConnectionString;
            //string connectionstringtxt = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;
            SqlCommand cmd = new SqlCommand();
            //SqlConnection conn = new SqlConnection();

            try
            {
                //truncate daily table
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    //cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "truncate table dbo.tbl_outlook_mails_mailbox_summary_dotnet_kycchecks_mumbai_mailbox_archive";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ab)
            {
                //Console.WriteLine("Error Generated Details: " + ab.ToString());
            }

            //DateTime startdate = new DateTime(2025, 08, 01);
            //DateTime enddate = new DateTime(2025, 08, 20);

            // Get today's date
            DateTime today = DateTime.Today;
            // Calculate the date 3 days ago
            DateTime threeDaysAgo = today.AddDays(-60);

            // Format the dates as "dd/MM/yyyy"
            string startDate = threeDaysAgo.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            string endDate = today.ToString("dd/MM/yyyy 23:59", CultureInfo.InvariantCulture);

            Outlook.Items filteredItems = null;

            

            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
                // Access the specific mailbox
                Outlook.MAPIFolder mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                Outlook.MAPIFolder inbox = mailbox.Folders["Inbox"];

                //Outlook.MAPIFolder inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.Items items = inbox.Items;

                //string filter = "[ReceivedTime] >= '01/08/2025' AND [ReceivedTime] <= '13/08/2025'";
                // Construct the filter string
                string filter = string.Format("[ReceivedTime] >= '{0}' AND [ReceivedTime] <= '{1}'", startDate, endDate);

                filteredItems = inbox.Items.Restrict(filter);
                filteredItems.Sort("[ReceivedTime]", true);
                //items.Sort("[ReceivedTime]", true);
                //string filter = $"[ReceivedTime] >= '{startdate:g}' AND [ReceivedTime] <= '{enddate:g}'";
                //Outlook.Items filteredItems = inbox.Items.Restrict(filter);

                //items.Sort("[ReceivedTime]", true);

                //get email address for people marked in To
                   


                foreach (object item in filteredItems)
                {
                    Outlook.MailItem mail = item as Outlook.MailItem;
                    if (mail != null)
                    {
                        try
                        {
                            DateTime receivedtime = mail.ReceivedTime;
                            string subject = mail.Subject;
                            //string body = mail.Body;
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
                                // The correct property for MailItem flag status is FlagStatus
                                Microsoft.Office.Interop.Outlook.OlFlagStatus olFlagStatus = mail.FlagStatus;
                                flagStatus = Enum.GetName(typeof(Microsoft.Office.Interop.Outlook.OlFlagStatus), olFlagStatus);
                            }
                            catch
                            {
                                // Handle cases where the status is not a standard value
                                flagStatus = "Unknown";
                            }

                            /*
                            // Get sender email address properly
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
                            catch (Exception ex)
                            {
                                // Optional: log or handle the exception
                                senderEmail = mail.SenderEmailAddress; // fallback
                            }
                             */

                            using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                            {
                                conn.Open();
                                cmd.Parameters.Clear();
                                cmd.Connection = conn;
                                //cmd.CommandType = CommandType.Text;
                                cmd.CommandText = "INSERT INTO dbo.tbl_outlook_mails_mailbox_summary_dotnet_kycchecks_mumbai_mailbox_archive (Subject,ReceivedDateTime,Sender,Categories,CC,Importance,EntryID,UploadDateTime,IsUnread,[TO],IsFlagged,FlagStatus) VALUES (@Subject,@ReceivedDateTime,@Sender,@Categories,@CC,@Importance,@EntryID,@UploadDateTime,@IsUnread,@TO,@IsFlagged,@FlagStatus)";
                                cmd.Parameters.AddWithValue("@Subject", subject ?? "");
                                //cmd.Parameters.AddWithValue("@Body", body ?? "");
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
                            // Release COM object for individual mail item
                            Marshal.ReleaseComObject(mail);

                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mail);
                            Marshal.ReleaseComObject(filteredItems);
                            Marshal.ReleaseComObject(items);
                            Marshal.ReleaseComObject(inbox);
                            Marshal.ReleaseComObject(mailbox);
                            Marshal.ReleaseComObject(outlookNamespace);
                            Marshal.ReleaseComObject(outlookApp);
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                            //if (mail != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);

                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(filteredItems);
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(mailbox);
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNamespace);
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);

                        }
                    }
                    Marshal.ReleaseComObject(item);
                    Marshal.ReleaseComObject(filteredItems);
                    Marshal.ReleaseComObject(items);
                    Marshal.ReleaseComObject(inbox);
                    Marshal.ReleaseComObject(mailbox);
                    Marshal.ReleaseComObject(outlookNamespace);
                    Marshal.ReleaseComObject(outlookApp);
                }
            


                //MessageBox.Show("Emails successfully saved to database");
            }
            catch (Exception ab)
            {
                //Console.WriteLine("Error Generated Details: " + ab.ToString());
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
                    cmd.CommandText = "exec dbo.usp_reading_mails_outlook_dotnet_kycchecks_mumbai_mailbox_summary";
                    //MessageBox.Show("Emails successfully saved to database");
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ab)
            {
                //Console.WriteLine("Error Generated Details: " + ab.ToString());
            }
             
        }
    }
}
