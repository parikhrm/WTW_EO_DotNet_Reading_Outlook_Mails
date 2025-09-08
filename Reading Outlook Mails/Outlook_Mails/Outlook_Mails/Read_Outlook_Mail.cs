using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Configuration;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using System.Windows.Forms.Integration;
using System.Windows.Forms.Design;
using System.Runtime.InteropServices;
using System.Globalization;

namespace Outlook_Mails
{
    public partial class Read_Outlook_Mail : Form
    {
        public string connectionstringtxt = "Data Source=A20-CB-DBSE01P;Initial Catalog=DRD;User ID=DRDUsers;Password=24252425";
        //public string connectionstringtxt = ConfigurationManager.ConnectionStrings["KYC_RDC_Workflow.Properties.Settings.DRDConnectionString"].ConnectionString;
        //string connectionstringtxt = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;
        SqlCommand cmd = new SqlCommand();
        SqlConnection conn = new SqlConnection();

        public Read_Outlook_Mail()
        {
            InitializeComponent();
        }

        private void Read_Outlook_Mail_Load(object sender, EventArgs e)
        {
            //reading_mails();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            reading_mails();
            //reading_mails_bulkinsert();
            //reading_mails_google();
            //reading_mails_bulkinsert_google();
            
        }

        public void reading_mails_google()
        {
            DateTime startdate = new DateTime(2025, 08, 01);
            DateTime enddate = new DateTime(2025, 08, 02, 23, 59, 59); // Include the entire end date

            Outlook.Application outlookApp = null;
            Outlook.NameSpace outlookNamespace = null;
            Outlook.MAPIFolder mailbox = null;
            Outlook.MAPIFolder inbox = null;
            Outlook.Items items = null;
            Outlook.Items filteredItems = null;

            try
            {
                outlookApp = new Outlook.Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
                mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                inbox = mailbox.Folders["Inbox"];
                items = inbox.Items;

                // Construct the filter string dynamically using the startdate and enddate variables
                //  Use the "g" format specifier to get a general date/time pattern
                //string filter = "[ReceivedTime] >= '{startdate:g}' AND [ReceivedTime] <= '{enddate:g}'";
                string filter = "[ReceivedTime] >= '01/08/2025' AND [ReceivedTime] <= '05/08/2025'";
                filteredItems = items.Restrict(filter); // Apply the filter to the items collection

                // Sorting is done on the filtered collection to avoid re-sorting the entire inbox items
                filteredItems.Sort("[ReceivedTime]", true);

                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();

                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO dbo.tbl_outlook_mails_dotnet (Subject,ReceivedDateTime,Sender,Categories,CC,Importance,EntryID,UploadDateTime,IsUnread) VALUES (@Subject,@ReceivedDateTime,@Sender,@Categories,@CC,@Importance,@EntryID,@UploadDateTime,@IsUnread)";

                    // Use a for loop to iterate through the filteredItems collection
                    // This avoids the issue of implicit enumerators and helps manage COM objects more effectively
                    for (int i = 1; i <= filteredItems.Count; i++) // Outlook collections are 1-based
                    {
                        Outlook.MailItem mail = null;
                        Outlook.AddressEntry senderEntry = null;
                        Outlook.ExchangeUser exchUser = null;

                        try
                        {
                            mail = filteredItems[i] as Outlook.MailItem; // Access item by index

                            if (mail != null)
                            {
                                // The date check within the loop is now redundant due to the Restrict filter,
                                // but included for demonstrative purposes and to show how to check date properties.
                                // if (mail.ReceivedTime >= startdate && mail.ReceivedTime <= enddate) 
                                //  {
                                string subject = mail.Subject;
                                string cc = mail.CC;
                                string categories = mail.Categories;
                                var importance = mail.Importance;
                                string entryid = mail.EntryID;
                                bool isunread = mail.UnRead;

                                string senderEmail = string.Empty;

                                try
                                {
                                    if (mail.SenderEmailType == "EX")
                                    {
                                        senderEntry = mail.Sender;
                                        if (senderEntry != null)
                                        {
                                            exchUser = senderEntry.GetExchangeUser();
                                            if (exchUser != null && !string.IsNullOrEmpty(exchUser.PrimarySmtpAddress))
                                            {
                                                senderEmail = exchUser.PrimarySmtpAddress;
                                            }
                                            else
                                            {
                                                senderEmail = senderEntry.Address;
                                            }
                                        }
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
                                finally
                                {
                                    // Release COM objects as soon as they are no longer needed
                                    if (exchUser != null) Marshal.ReleaseComObject(exchUser);
                                    if (senderEntry != null) Marshal.ReleaseComObject(senderEntry);
                                }

                                // Parameterized query helps prevent SQL Injection attacks
                                cmd.Parameters.Clear();
                                cmd.Parameters.AddWithValue("@Subject", subject ?? "");
                                cmd.Parameters.AddWithValue("@ReceivedDateTime", mail.ReceivedTime);
                                cmd.Parameters.AddWithValue("@Sender", senderEmail ?? "");
                                cmd.Parameters.AddWithValue("@Categories", categories ?? "");
                                cmd.Parameters.AddWithValue("@CC", cc ?? "");
                                cmd.Parameters.AddWithValue("@Importance", importance);
                                cmd.Parameters.AddWithValue("@EntryID", entryid ?? "");
                                cmd.Parameters.AddWithValue("@UploadDateTime", DateTime.Now.ToLocalTime());
                                cmd.Parameters.AddWithValue("@IsUnread", isunread);
                                cmd.ExecuteNonQuery();
                                // }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log mail processing errors here if needed
                            MessageBox.Show("Error Generated Details: " + ex.ToString());
                        }
                        finally
                        {
                            if (mail != null) Marshal.ReleaseComObject(mail); // Release mail item within the loop
                        }
                    }
                } // SqlConnection is automatically closed and disposed when the using block exits
                MessageBox.Show("Emails successfully saved to database");
            }
            catch (Exception ab)
            {
                MessageBox.Show("Error Generated Details: " + ab.ToString());
            }
            finally
            {
                // Release COM objects in reverse order of creation, outside the loop
                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                if (items != null) Marshal.ReleaseComObject(items);
                if (inbox != null) Marshal.ReleaseComObject(inbox);
                if (mailbox != null) Marshal.ReleaseComObject(mailbox);
                if (outlookNamespace != null) Marshal.ReleaseComObject(outlookNamespace);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
        }

        public void reading_mails()
        {
            try
            {
                //truncate daily table
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "truncate table dbo.tbl_outlook_mails_daily_dotnet";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ab)
            {
                MessageBox.Show("Error Generated Details: " + ab.ToString());
            }
            
            
            //DateTime startdate = new DateTime(2025, 08, 01);
            //DateTime enddate = new DateTime(2025, 08, 20);

            // Get today's date
            DateTime today = DateTime.Today;
            // Calculate the date 3 days ago
            DateTime threeDaysAgo = today.AddDays(-2);

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
                            //string sender = mail.SenderEmailAddress;
                            string cc = mail.CC;
                            string categories = mail.Categories;
                            var importance = mail.Importance;
                            string entryid = mail.EntryID;
                            bool isunread = mail.UnRead;

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

                            using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                            {
                                conn.Open();
                                cmd.Parameters.Clear();
                                cmd.Connection = conn;
                                cmd.CommandType = CommandType.Text;
                                cmd.CommandText = "INSERT INTO dbo.tbl_outlook_mails_daily_dotnet (Subject,ReceivedDateTime,Sender,Categories,CC,Importance,EntryID,UploadDateTime,IsUnread) VALUES (@Subject,@ReceivedDateTime,@Sender,@Categories,@CC,@Importance,@EntryID,@UploadDateTime,@IsUnread)";
                                cmd.Parameters.AddWithValue("@Subject", subject ?? "");
                                //cmd.Parameters.AddWithValue("@Body", body ?? "");
                                cmd.Parameters.AddWithValue("@ReceivedDateTime", receivedtime);
                                cmd.Parameters.AddWithValue("@Sender", senderEmail ?? "");
                                cmd.Parameters.AddWithValue("@Categories", categories ?? "");
                                cmd.Parameters.AddWithValue("@CC", cc ?? "");
                                cmd.Parameters.AddWithValue("@Importance", importance);
                                cmd.Parameters.AddWithValue("@EntryID", entryid ?? "");
                                cmd.Parameters.AddWithValue("@UploadDateTime", DateTime.Now.ToLocalTime());
                                cmd.Parameters.AddWithValue("@IsUnread", isunread);
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
                MessageBox.Show("Error Generated Details: " + ab.ToString());
            }

            try
            {
                //run stored procedure
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "dbo.usp_reading_mails_outlook_dotnet";
                    MessageBox.Show("Emails successfully saved to database");
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ab)
            {
                MessageBox.Show("Error Generated Details: " + ab.ToString());
            }
        }

        public void reading_mails_bulkinsert()
        {
            
            DateTime startdate = new DateTime(2025, 08, 01);
            DateTime enddate = new DateTime(2025, 08, 20);

            Outlook.Items filteredItems = null;

            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                // Access the specific mailbox
                Outlook.MAPIFolder mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                Outlook.MAPIFolder inbox = mailbox.Folders["Inbox"];
                Outlook.Items items = inbox.Items;

                string filter = "[ReceivedTime] >= '01/08/2025' AND [ReceivedTime] <= '02/08/2025'";
                filteredItems = items.Restrict(filter); // Apply the filter to the items collection

                //items.Sort("[ReceivedTime]", true);
                filteredItems.Sort("[ReceivedTime]", true);

                // Prepare DataTable
                DataTable emailTable = new DataTable();
                emailTable.Columns.Add("Subject", typeof(string));
                emailTable.Columns.Add("ReceivedDateTime", typeof(DateTime));
                emailTable.Columns.Add("Sender", typeof(string));
                emailTable.Columns.Add("Categories", typeof(string));
                emailTable.Columns.Add("CC", typeof(string));
                emailTable.Columns.Add("Importance", typeof(int)); // Enum stored as int
                emailTable.Columns.Add("EntryID", typeof(string));
                emailTable.Columns.Add("UploadDateTime", typeof(DateTime));

                for (int i = 1; i <= filteredItems.Count; i++)
                {
                    object item = filteredItems[i];
                    Outlook.MailItem mail = item as Outlook.MailItem;
                    //Outlook.MailItem mail = null;
                    Outlook.AddressEntry senderEntry = null;
                    Outlook.ExchangeUser exchUser = null;

                    if (mail != null)
                    {
                        try
                        {
                            DateTime receivedtime = mail.ReceivedTime;

                            string senderEmail = string.Empty;

                            try
                            {
                                if (mail.SenderEmailType == "EX")
                                {
                                    senderEntry = mail.Sender;
                                    if (senderEntry != null)
                                    {
                                        exchUser = senderEntry.GetExchangeUser();
                                        if (exchUser != null && !string.IsNullOrEmpty(exchUser.PrimarySmtpAddress))
                                        {
                                            senderEmail = exchUser.PrimarySmtpAddress;
                                        }
                                        else
                                        {
                                            senderEmail = senderEntry.Address;
                                        }
                                    }
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
                            finally
                            {
                                // Release COM objects as soon as they are no longer needed
                                if (exchUser != null) Marshal.ReleaseComObject(exchUser);
                                if (senderEntry != null) Marshal.ReleaseComObject(senderEntry);
                            }

                            emailTable.Rows.Add(
                            mail.Subject ?? "",
                            receivedtime,
                                //mail.SenderEmailAddress ?? "",
                            senderEmail,
                            mail.Categories ?? "",
                            mail.CC ?? "",
                            (int)mail.Importance,
                            mail.EntryID ?? "",
                            DateTime.Now.ToLocalTime()
                            );

                        }
                        finally
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                        }
                    }
                    if (item != null) Marshal.ReleaseComObject(item);
                }

                // Bulk insert into SQL Server
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.TableLock, null))
                    {
                        bulkCopy.DestinationTableName = "dbo.tbl_outlook_mails_dotnet";
                        bulkCopy.BatchSize = 500;
                        bulkCopy.BulkCopyTimeout = 600;

                        bulkCopy.ColumnMappings.Add("Subject", "Subject");
                        bulkCopy.ColumnMappings.Add("ReceivedDateTime", "ReceivedDateTime");
                        bulkCopy.ColumnMappings.Add("Sender", "Sender");
                        bulkCopy.ColumnMappings.Add("Categories", "Categories");
                        bulkCopy.ColumnMappings.Add("CC", "CC");
                        bulkCopy.ColumnMappings.Add("Importance", "Importance");
                        bulkCopy.ColumnMappings.Add("EntryID", "EntryID");
                        bulkCopy.ColumnMappings.Add("UploadDateTime", "UploadDateTime");

                        bulkCopy.WriteToServer(emailTable);
                    }
                }

                MessageBox.Show("Emails successfully saved to database");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Generated Details: " + ex.ToString());
            }
            finally
            {
                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                //if (items != null) Marshal.ReleaseComObject(items); //
                //if (inbox != null) Marshal.ReleaseComObject(inbox); //
                //if (mailbox != null) Marshal.ReleaseComObject(mailbox); //
                //if (outlookNamespace != null) Marshal.ReleaseComObject(outlookNamespace); //
                //if (outlookApp != null) Marshal.ReleaseComObject(outlookApp); //
            }

        }

        public void reading_mails_bulkinsert_google()
        {
            DateTime startDate = new DateTime(2025, 08, 01);
            DateTime endDate = new DateTime(2025, 08, 20);

            Outlook.Application outlookApp = null;
            Outlook.NameSpace outlookNamespace = null;
            Outlook.MAPIFolder mailbox = null;
            Outlook.MAPIFolder inbox = null;
            Outlook.Items items = null;
            Outlook.Items filteredItems = null;

            try
            {
                // 1. Initialize Outlook objects at the last possible moment
                outlookApp = new Outlook.Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");

                // Access the specific mailbox
                mailbox = outlookNamespace.Folders["kycchecksmumbai"];
                inbox = mailbox.Folders["Inbox"];
                items = inbox.Items;

                // 2. Dynamically construct the Restrict filter string using parameters
                //string filter = $"[ReceivedTime] >= '{startDate.ToString("MM/dd/yyyy")}' AND [ReceivedTime] <= '{endDate.ToString("MM/dd/yyyy")}'"; //
                string filter = "[ReceivedTime] >= '01/08/2025' AND [ReceivedTime] <= '02/08/2025'";
                filteredItems = items.Restrict(filter); // Apply the filter to the items collection

                filteredItems.Sort("[ReceivedTime]", true);

                // Prepare DataTable
                DataTable emailTable = new DataTable();
                emailTable.Columns.Add("Subject", typeof(string));
                emailTable.Columns.Add("ReceivedDateTime", typeof(DateTime));
                emailTable.Columns.Add("Sender", typeof(string));
                emailTable.Columns.Add("Categories", typeof(string));
                emailTable.Columns.Add("CC", typeof(string));
                emailTable.Columns.Add("Importance", typeof(int));
                emailTable.Columns.Add("EntryID", typeof(string));
                emailTable.Columns.Add("UploadDateTime", typeof(DateTime));

                // Use BeginLoadData to optimize DataTable population
                emailTable.BeginLoadData(); //

                // 3. Avoid foreach loop for COM collections to prevent memory leaks
                //    Instead, use a for loop and release each MailItem after use.
                for (int i = 1; i <= filteredItems.Count; i++) // Outlook collection is 1-based
                {
                    object item = filteredItems[i]; // Get the item at the current index
                    Outlook.MailItem mail = item as Outlook.MailItem;
                    Outlook.AddressEntry senderEntry = null;
                    Outlook.ExchangeUser exchUser = null;

                    try
                    {
                        mail = filteredItems[i] as Outlook.MailItem;
                        if (mail != null)
                        {
                            try
                            {
                                string senderEmail = string.Empty;
                                try
                                {
                                    // 4. Efficient sender email resolution
                                    if (mail.SenderEmailType == "EX")
                                    {
                                        senderEntry = mail.Sender;
                                        if (senderEntry != null)
                                        {
                                            exchUser = senderEntry.GetExchangeUser();
                                            if (exchUser != null && !string.IsNullOrEmpty(exchUser.PrimarySmtpAddress)) //
                                            {
                                                senderEmail = exchUser.PrimarySmtpAddress; //
                                            }
                                            else
                                            {
                                                senderEmail = senderEntry.Address;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        senderEmail = mail.SenderEmailAddress; //
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Optional: log or handle the exception (e.g., if Active Directory resolution fails)
                                    senderEmail = mail.SenderEmailAddress; // fallback
                                    // Log.Warn($"Error resolving sender email for mail item with Subject: {mail.Subject}. Falling back to SenderEmailAddress. Error: {ex.Message}");
                                }
                                finally
                                {
                                    // Release COM objects as soon as they are no longer needed
                                    if (exchUser != null) Marshal.ReleaseComObject(exchUser); //
                                    if (senderEntry != null) Marshal.ReleaseComObject(senderEntry); //
                                }

                                emailTable.Rows.Add(
                                    mail.Subject ?? "",
                                    mail.ReceivedTime,
                                    senderEmail,
                                    mail.Categories ?? "",
                                    mail.CC ?? "",
                                    (int)mail.Importance,
                                    mail.EntryID ?? "",
                                    DateTime.Now.ToLocalTime()
                                );

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error processing mail item: " + ex.ToString());
                                // Log the exception for better debugging
                                // Log.Error("Error processing mail item", ex);
                            }
                            finally
                            {
                                // Release the mail object immediately after use
                                Marshal.ReleaseComObject(mail); //
                            }
                        }
                        if (item != null) Marshal.ReleaseComObject(item); // Release the generic item object too
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error processing mail item: " + ex.ToString());
                    }
                }

                // Bulk insert into SQL Server
                string connectionstringtxt = "your_connection_string"; // Replace with your actual connection string
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.TableLock, null)) // Using TableLock for potentially better performance
                    {
                        bulkCopy.DestinationTableName = "dbo.tbl_outlook_mails_dotnet";
                        bulkCopy.BatchSize = 500; // Consider adjusting for optimal performance
                        bulkCopy.BulkCopyTimeout = 600; // Increase timeout if needed for large transfers

                        // Column mappings for SqlBulkCopy
                        bulkCopy.ColumnMappings.Add("Subject", "Subject");
                        bulkCopy.ColumnMappings.Add("ReceivedDateTime", "ReceivedDateTime");
                        bulkCopy.ColumnMappings.Add("Sender", "Sender");
                        bulkCopy.ColumnMappings.Add("Categories", "Categories");
                        bulkCopy.ColumnMappings.Add("CC", "CC");
                        bulkCopy.ColumnMappings.Add("Importance", "Importance");
                        bulkCopy.ColumnMappings.Add("EntryID", "EntryID");
                        bulkCopy.ColumnMappings.Add("UploadDateTime", "UploadDateTime");

                        bulkCopy.WriteToServer(emailTable);
                    }
                }

                MessageBox.Show("Emails successfully saved to database");
            }
            catch (Exception ex)
            {
                // 5. Improved Error Handling and Logging
                MessageBox.Show("Error Generated Details: " + ex.ToString());
                // Consider using a dedicated logging framework (e.g., NLog, Serilog)
                // Log.Error("Outlook processing and data insertion failed", ex);
            }
            finally
            {
                // Release Outlook COM objects in reverse order of creation
                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems); //
                if (items != null) Marshal.ReleaseComObject(items); //
                if (inbox != null) Marshal.ReleaseComObject(inbox); //
                if (mailbox != null) Marshal.ReleaseComObject(mailbox); //
                if (outlookNamespace != null) Marshal.ReleaseComObject(outlookNamespace); //
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp); //
            }
        }
        

    }

}   
