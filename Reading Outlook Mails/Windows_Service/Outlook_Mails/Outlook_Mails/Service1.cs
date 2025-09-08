using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
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


namespace Outlook_Mails
{
    public partial class Service1 : ServiceBase
    {
        private string connectionstringtxt = "Data Source=A20-CB-DBSE01P;Initial Catalog=DRD;User ID=DRDUsers;Password=24252425";
        //public string connectionstringtxt = ConfigurationManager.ConnectionStrings["KYC_RDC_Workflow.Properties.Settings.DRDConnectionString"].ConnectionString;
        //string connectionstringtxt = System.Configuration.ConfigurationManager.ConnectionStrings["connection_string"].ConnectionString;
        //SqlCommand cmd = new SqlCommand();
        SqlConnection conn = new SqlConnection();
        private SqlCommand cmd = new SqlCommand(); // Shared command object

        private Timer timer;
        private const double IntervalMilliseconds = 30 * 60 * 1000; // 30 minutes in milliseconds

        public Service1()
        {
            InitializeComponent();
            ServiceName = "OutlookMailProcessorService"; 
        }

        protected override void OnStart(string[] args)
        {
            // Initialize the timer
            timer = new Timer(IntervalMilliseconds); // {Link: According to ironpdf.com https://ironpdf.com/blog/net-help/csharp-timer/}
            timer.Elapsed += Timer_Elapsed; // Attach the event handler
            timer.AutoReset = true; // Set to true to repeat the timer
            timer.Enabled = true; // Start the timer

            // Optional: Log service start
            WriteToEventLog("Service started.");

            // Optionally, you can trigger the processing immediately on service start
            // Timer_Elapsed(null, null);
        }

        protected override void OnStop()
        {
            if (timer != null)
            {
                timer.Stop();
                timer.Dispose();
            }

            // Optional: Log service stop
            WriteToEventLog("Service stopped.");
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            // The core logic that will run every 30 minutes
            ProcessOutlookMails();
        }

        private void ProcessOutlookMails()
        {
            try
            {
                // Truncate daily table
                using (SqlConnection conn = new SqlConnection(connectionstringtxt))
                {
                    conn.Open();
                    cmd.Parameters.Clear();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "truncate table dbo.tbl_outlook_mails_daily_dotnet";
                    cmd.ExecuteNonQuery();
                }

                // Get today's date
                DateTime today = DateTime.Today;
                // Calculate the date 2 days ago (as per your original code)
                DateTime threeDaysAgo = today.AddDays(-2);

                // Format the dates as "dd/MM/yyyy"
                string startDate = threeDaysAgo.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                string endDate = today.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

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

                    string filter = string.Format("[ReceivedTime] >= '{0}' AND [ReceivedTime] <= '{1}'", startDate, endDate);
                    filteredItems = inbox.Items.Restrict(filter);
                    filteredItems.Sort("[ReceivedTime]", true);

                    foreach (object item in filteredItems)
                    {
                        Outlook.MailItem mail = item as Outlook.MailItem;
                        if (mail != null)
                        {
                            string subject = mail.Subject;
                            DateTime receivedtime = mail.ReceivedTime;
                            string cc = mail.CC;
                            string categories = mail.Categories;
                            Outlook.OlImportance importance = mail.Importance; // Explicitly use Outlook.OlImportance
                            string entryid = mail.EntryID;
                            bool isunread = mail.UnRead;

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
                                WriteToEventLog("Error getting sender email: " + ex.Message);
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
                    }
                }
                catch (Exception ex)
                {
                    WriteToEventLog("Error during Outlook processing: " + ex.Message + " Stack Trace: " + ex.StackTrace);
                }
                finally
                {
                    // Ensure all COM objects are released
                    if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                    if (items != null) Marshal.ReleaseComObject(items); // items is not used, but included for completeness
                    if (inbox != null) Marshal.ReleaseComObject(inbox);
                    if (mailbox != null) Marshal.ReleaseComObject(mailbox);
                    if (outlookNamespace != null) Marshal.ReleaseComObject(outlookNamespace);
                    if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
                }
            }
            catch (Exception serviceEx)
            {
                // Catching exceptions from the database operations or higher level logic
                WriteToEventLog("Error in ProcessOutlookMails: " + serviceEx.Message + " Stack Trace: " + serviceEx.StackTrace);
            }
        }

        private void WriteToEventLog(string message)
        {
            // For a Windows Service, you typically write to the Windows Event Log
            // or a custom log file instead of using MessageBox.Show.
            // Make sure your service account has permission to write to the event log.
            using (System.Diagnostics.EventLog eventLog = new System.Diagnostics.EventLog("Application"))
            {
                eventLog.Source = "OutlookMailProcessorService"; // Replace with your service name
                eventLog.WriteEntry(message, System.Diagnostics.EventLogEntryType.Information);
            }
        }

        


    }
}
