using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Net.Mail;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("Simple Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("Simple Service stopped {0}");
            this.Schedular.Dispose();
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
                this.WriteToFile("Simple Service Mode: " + mode + " {0}");

                

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                    }
                }

                if (mode.ToUpper() == "INTERVAL")
                {
                    //Get the Interval in Minutes from AppSettings.
                    int intervalMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"]);

                    //Set the Scheduled Time by adding the Interval to Current Time.
                    scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                        scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                    }
                }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                this.WriteToFile("Simple Service scheduled to run after: " + schedule + " {0}");

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("Baza Przypomnienie"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void SchedularCallback(object e)
        {

            try
            {
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                string query2 = "SELECT * FROM xUsers WHERE Email IS NOT NULL";
                string query3 = "SELECT ID_uzytkownika, Data, (SELECT Czy_roboczy FROM Kalendarz As k WHERE (DATEPART (DAY, k.Data) =@Day AND DATEPART (Month, k.Data) =@Month AND DATEPART (Year, k.Data) =@Year)) AS czy_rob FROM dPracePrzyProjektach AS p WHERE DATEPART (DAY, p.Data) =@Day AND DATEPART (Month, p.Data) =@Month AND DATEPART (Year, p.Data) =@Year";
                string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;

                SqlConnection conn = new SqlConnection(constr);
                try
                {
                    conn.Open();
                }
                catch (Exception ex)
                {
                    WriteToFile(ex.Message);
                }


                SqlDataAdapter da = new SqlDataAdapter(query2, constr);
                da.Fill(ds, "xUsers");

                using (SqlCommand cmd = new SqlCommand(query3))
                {
                    cmd.Connection = conn;
                    cmd.Parameters.AddWithValue("@Day", DateTime.Today.Day - 1);
                    cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month);
                    cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds, "dPracePrzyProjektach");
                    }
                }
                bool rob = (from DataRow dr in ds.Tables["dPracePrzyProjektach"].Rows
                            select (bool)dr["czy_rob"]).FirstOrDefault();
                if (rob)
                {

                    bool[] wierszeDoUsuniecia = new bool[ds.Tables["xUsers"].Rows.Count];

                    for (int j = 0; j < ds.Tables["xUsers"].Rows.Count; j++)
                    {
                        DataRow row = ds.Tables["xUsers"].Rows[j];
                        string ID = row["ID"].ToString();

                        for (int i = ds.Tables["dPracePrzyProjektach"].Rows.Count - 1; i >= 0; i--)
                        {
                            DataRow delIDS = ds.Tables["dPracePrzyProjektach"].Rows[i];
                            if (delIDS["ID_uzytkownika"].ToString() == ID)
                            {
                                try
                                {
                                    wierszeDoUsuniecia[j] = true;
                                }
                                catch { }
                            }
                        }

                    }

                    for (int i = ds.Tables["xUsers"].Rows.Count - 1; i >= 0; i--)
                    {
                        if (wierszeDoUsuniecia[i])
                        {
                            DataRow drrr = ds.Tables["xUsers"].Rows[i];
                            ds.Tables["xUsers"].Rows.Remove(drrr);
                        }
                    }

                    foreach (DataRow row in ds.Tables["xUsers"].Rows)
                    {


                        string email = row["Email"].ToString();
                        //string data = row["Data"].ToString();
                        string imie = row["Imie"].ToString();

                        WriteToFile("Trying to send email to: " + email);

                        using (MailMessage mm = new MailMessage("Unimap.katowice@gmail.com", email))
                        {
                            mm.Subject = string.Format("Brak wpisu w bazie!");
                            mm.Body = string.Format("Witaj "+ imie + "!<br></br> <br></br> " + "Proszę o uzupełnienie bazy za wczorajszy dzień <br></br> <br></br><a href=http://unimap.homenet.org:8080/unimap_UniWorker/startPage.aspx?Opis=Witamy%20na%20stronie%20rejestracji%20przebiegu%20projekt%C3%B3w> link do bazy</a>");

                            mm.IsBodyHtml = true;
                            SmtpClient smtp = new SmtpClient();
                            smtp.Host = "smtp.gmail.com";
                            smtp.EnableSsl = true;
                            System.Net.NetworkCredential credentials = new System.Net.NetworkCredential();
                            credentials.UserName = "Unimap.katowice@gmail.com";
                            credentials.Password = "MalBry2015";
                            smtp.UseDefaultCredentials = true;
                            smtp.Credentials = credentials;
                            smtp.Port = 587;
                            smtp.Send(mm);
                            WriteToFile("Email sent successfully to: " + email);
                        }
                    } 
                }
                else
                {
                    WriteToFile("Dzien wolny");
                }
            this.ScheduleService();
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);
                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("Baza Przypomnienie"))
                {
                    serviceController.Stop();
                }
            }

        }

        private void WriteToFile(string text)
        {
            string path = "C:\\BazaPrzypomnienie_ServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }
    }
}
