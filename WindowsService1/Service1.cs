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
                //ustawianie dnia raportowania
                string RaportDay = ConfigurationManager.AppSettings["RaportDay"].ToUpper();
                this.WriteToFile("Simple Service Mode: " + RaportDay + " {0}");

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

                /////wysłanie maili z brakami w raporcie
                if (DateTime.Today.Day == Convert.ToInt32(RaportDay))
                {
                    string query4 = "SELECT Count(czy_roboczy)*8 AS m_godz FROM Kalendarz AS k WHERE DATEPART (yyyy, k.Data) = @Year AND DATEPART (m, k.Data) = @Month AND czy_roboczy='TRUE'";
                    string query5 = "SELECT ID_uzytkownika, Sum(Datediff(ss, Godzina_od, Godzina_do) / 3600.0) AS m_rob FROM dPracePrzyProjektach WHERE Datepart(m, Data) = @Month AND Datepart(yyyy, Data) = @Year  Group BY ID_uzytkownika";
                    string query6 = "SELECT ID_uzytkownika, Data, Sum(Datediff(mi, Godzina_od, Godzina_do) / 60.00) AS d_rob FROM dPracePrzyProjektach WHERE Datepart(m, Data) = @Month AND Datepart(yyyy, Data) = @Year Group BY ID_uzytkownika, Data";
                    //string query6 = "SELECT dPracePrzyProjektach.ID_uzytkownika, Kalendarz.Data, Kalendarz.Dzien_tygodnia, Kalendarz.Czy_roboczy, Sum(Datediff(ss, dPracePrzyProjektach.Godzina_od, dPracePrzyProjektach.Godzina_do) / 3600.0) AS d_rob FROM Kalendarz LEFT OUTER JOIN dPracePrzyProjektach ON dPracePrzyProjektach.Data=Kalendarz.Data WHERE Datepart(m, Kalendarz.Data) = 5 AND Datepart(yyyy, Kalendarz.Data) = 2017 Group BY dPracePrzyProjektach.ID_uzytkownika, Kalendarz.Data, Kalendarz.Dzien_tygodnia, Kalendarz.Czy_roboczy ORDER BY Kalendarz.Data";
                    string query7 = "SELECT Data, Dzien_tygodnia, Czy_roboczy FROM Kalendarz WHERE Datepart(m, Data) = @Month AND Datepart(yyyy, Data) = @Year";

                    using (SqlCommand cmd = new SqlCommand(query5))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month - 1);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda = new SqlDataAdapter(cmd))
                        {
                            mda.Fill(ds, "rob_miesiac");
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand(query4))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month - 1);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda2 = new SqlDataAdapter(cmd))
                        {
                            mda2.Fill(ds, "miesiac");
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand(query6))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month - 1);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda = new SqlDataAdapter(cmd))
                        {
                            mda.Fill(ds, "rob_dni");
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand(query7))
                    {
                        cmd.Connection = conn;
                        cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month - 1);
                        cmd.Parameters.AddWithValue("@Year", DateTime.Today.Year);
                        using (SqlDataAdapter mda = new SqlDataAdapter(cmd))
                        {
                            mda.Fill(ds, "kalend");
                        }
                    }
                    foreach (DataRow row in ds.Tables["xUsers"].Rows)
                    {
                        string messageBody = "<font>Poniżej zamieszczono wykaz z bazy za poprzedni miesiąc: </font><br><br>";
                        string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                        string htmlTableEnd = "</table>";
                        string htmlHeaderRowStart = "<tr style =\"background-color:#6FA1D2; color:#ffffff;\">";
                        string htmlHeaderRowEnd = "</tr>";
                        string htmlTrStart = "<tr style =\"color:#555555;\">";
                        string htmlTrEnd = "</tr>";
                        string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                        string htmlTdStart2 = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;background-color:#D3D3D3; color:#ff0000;\">";
                        string htmlTdEnd = "</td>";

                        messageBody += htmlTableStart;
                        messageBody += htmlHeaderRowStart;
                        messageBody += htmlTdStart + "Data" + htmlTdEnd;
                        messageBody += htmlTdStart + "Dzień tygodnia" + htmlTdEnd;
                        messageBody += htmlTdStart + "Ilość przepracowanych godzin " + htmlTdEnd;
                        messageBody += htmlHeaderRowEnd;

                        foreach (DataRow Rowk in ds.Tables["kalend"].Rows)
                        {
                            messageBody = messageBody + htmlTrStart;
                            if ((Boolean)Rowk["Czy_roboczy"] == false)
                            {
                                string data = Rowk["Data"].ToString();
                                messageBody = messageBody + htmlTdStart2 + data + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart2 + Rowk["Dzien_tygodnia"] + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart2 + htmlTdEnd;
                            }
                            else
                            {
                                string data = Rowk["Data"].ToString();
                                messageBody = messageBody + htmlTdStart + data + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Rowk["Dzien_tygodnia"] + htmlTdEnd;
                            }

                            foreach (DataRow Rowr in ds.Tables["rob_dni"].Rows)
                            {
                                if (DateTime.Parse(Rowr["Data"].ToString()).ToShortDateString() == DateTime.Parse(Rowk["Data"].ToString()).ToShortDateString() & Rowr["Id_uzytkownika"].ToString() == row["ID"].ToString())
                                {
                                    messageBody = messageBody + htmlTdStart + Rowr["d_rob"] + htmlTdEnd;
                                    messageBody = messageBody + htmlTrEnd;
                                }
                            }
                        }
                        string tekst = "";
                        foreach (DataRow rowm in ds.Tables["rob_miesiac"].Rows)
                        {
                            if (rowm["Id_uzytkownika"].ToString() == row["ID"].ToString())
                            {
                                double m_rob = Convert.ToDouble(rowm["m_rob"]);
                                double m_godz = Convert.ToDouble(ds.Tables["miesiac"].Rows[0][0]);
                                if (m_rob > m_godz)
                                {
                                    double wydruk = Math.Round((m_rob - m_godz)*2)/ 2;
                                    tekst = "w poprzednim miesiącu  przepracowałeś/aś " + Math.Round(m_rob * 2) / 2 + " godzin w tym masz nadgodziny w liczbie " + wydruk + " godzin";
                                }
                                if (m_rob < m_godz)
                                {
                                    double wydruk = Math.Round((m_godz - m_rob)*2)/ 2;
                                    tekst = "w poprzednim miesiącu przepracowałeś/aś "+ Math.Round(m_rob * 2) / 2 + " godzin, ale masz za mało przepracowanych godzin o " + wydruk + " godziny";
                                }
                                if (m_rob == m_godz)
                                {
                                    tekst = "w poprzednim miesiącu masz przepracowane równe " + Math.Round(m_rob * 2) / 2 + " godziny";
                                }
                            }
                        }
                        messageBody = messageBody + htmlTableEnd + tekst;


                        string email = row["Email"].ToString();
                        //string email = "s3rek92@gmail.com";
                        if (email == row["Email"].ToString())
                        {
                            WriteToFile("Próba wysłania maila z raportem do: " + row["Email"]);

                            using (MailMessage mm = new MailMessage("Unimap.katowice@gmail.com", email))
                            {
                                mm.Subject = "Miesięczna kontrola bazy";
                                mm.IsBodyHtml = true;
                                mm.Body = string.Format("<p>Witaj! " + row["Imie"] + " " + row["Nazwisko"] + "</p>" + messageBody);

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
                                WriteToFile("Email z raportem wysłany do: " + email);
                            }
                        }
                    }
                }
                ///////////////////////////////////

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

                        //string email = "s3rek92@gmail.com";
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
