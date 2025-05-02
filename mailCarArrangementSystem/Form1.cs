using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.IO;
using System.Net.Mime;


namespace mailCarArrangementSystem
{
    public partial class Form1 : Form
    {
        bool start = false;
        int cnt = 0;
        string destinationPath;
        bool allowClose = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                start = true;
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    lblError.Text = ex.Message;
                });
            }
        }

        private void bgWork_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                DataTable dt = new DataTable();
                DataTable dtMail = new DataTable();
                dt = null;
                dtMail = null;
                string q = "SELECT ROW_NUMBER() OVER (ORDER BY PLAN_START_TIME ASC)   AS \"NO\", " +
                           "         TO_CHAR(A.PLAN_START_TIME, 'YYYY-MM-DD', 'NLS_DATE_LANGUAGE = ENGLISH') AS \"DATE\", " +
                           "         TO_CHAR(PLAN_START_TIME, 'HH24:MI') AS \"TIME\", " +
                           "         C.EXTRA_COLUMN1 AS PLAT_NO, " +
                           "         C.CODE_SHORT_NAME AS TYPE, " +
                           "         B.CODE_SHORT_NAME AS DRIVER, " +
                           "         B.EXTRA_COLUMN1 AS PHONE, " +
                           "         CASE WHEN A.RENTAL_DEPART_CD = 'OTH' THEN " +
                           "                  (SELECT MAX(CODE_SHORT_NAME)  " +
                           "                     FROM MSBS_CODE_MASTER@RJLMES " +
                           "                    WHERE CODE_CLASS_CD = 'RENTAL_CAR_DEPARTURE' " +
                           "                      AND CODE_NAME     = 'OTH') || ' - ' || A.RENTAL_DEPART_DESC " +
                           "             ELSE (SELECT MAX(CODE_SHORT_NAME)  " +
                           "                     FROM MSBS_CODE_MASTER@RJLMES " +
                           "                    WHERE CODE_CLASS_CD = 'RENTAL_CAR_DEPARTURE' " +
                           "                      AND CODE_NAME     = A.RENTAL_DEPART_CD) " +
                           "         END                                           AS DEPARTURE, " +
                           "         CASE WHEN A.RENTAL_PLACE_CD = 'OTH' THEN  " +
                           "                   (SELECT MAX(CODE_SHORT_NAME)  " +
                           "                      FROM MSBS_CODE_MASTER@RJLMES " +
                           "                     WHERE CODE_CLASS_CD = 'RENTAL_CAR_DESTINATION' " +
                           "                       AND CODE_NAME     = 'OTH') || ' - ' || A.RENTAL_PLACE_DESC " +
                           "              ELSE (SELECT MAX(CODE_SHORT_NAME)  " +
                           "                      FROM MSBS_CODE_MASTER@RJLMES " +
                           "                     WHERE CODE_CLASS_CD = 'RENTAL_CAR_DESTINATION' " +
                           "                       AND CODE_NAME     = A.RENTAL_PLACE_CD) " +
                           "          END                                           AS DESTINATION, " +
                           //"         (SELECT MAX(CODE_SHORT_NAME) " +
                           //"             FROM MSBS_CODE_MASTER " +
                           //"             WHERE CODE_CLASS_CD = 'RENTAL_CAR_DESTINATION' " +
                           //"             AND CODE_NAME     = A.RENTAL_PLACE_DESC) " +
                           //"         || ' - ' || UPPER(A.DATA_MEMO) AS DESTINATION, " +
                           "         UPPER(A.RENTAL_USED_DESC) AS NAME_OF_EMP " +
                           "     FROM SYS_RENTAL_DATA@RJLMES A " +
                           "     LEFT JOIN (SELECT * " +
                           "                 FROM MSBS_CODE_MASTER@RJLMES " +
                           "                 WHERE CODE_CLASS_CD = 'MASTER_DRIVER') B " +
                           "     ON A.EXTRA1_FLD = B.CODE_NAME  " +
                           "     LEFT JOIN (SELECT * " +
                           "                 FROM MSBS_CODE_MASTER@RJLMES " +
                           "                 WHERE CODE_CLASS_CD = 'MASTER_CAR') C " +
                           "     ON A.RENTAL_TYPE_CD = C.CODE_NAME   " +        
                           "     WHERE A.RENTAL_DIV = 'GA' " +
                           "     AND TO_CHAR(A.PLAN_START_TIME, 'YYYYMMDD') BETWEEN TO_CHAR(SYSDATE, 'YYYYMMDD')AND TO_CHAR(SYSDATE+1, 'YYYYMMDD') " +
                           "     AND A.RENTAL_STATUS = 'F' " +
                           "     ORDER BY A.PLAN_START_TIME ";
                dt = Class.CmdQry.getData(q);

                string q1 = "SELECT CODE_NAME     AS EMAIL, " +
                            "       EXTRA_COLUMN1 AS TYPE " +
                            "  FROM MSBS_CODE_MASTER@RJLMES " +
                            " WHERE CODE_CLASS_CD = 'MAIL_RENTAL_CAR_GROUPING' " +
                            "   AND USE_YN        = 'Y'";
                dtMail = Class.CmdQry.getData(q1);

                if (dt.Rows.Count > 0)
                {
                    if (dtMail.Rows.Count > 0)
                    {
                        fnSendMail(dt, dtMail);
                    }
                }

            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    lblError.Text = ex.Message;
                });
            }
        }

        private void fnSendMail(DataTable dt, DataTable dtMail)
        {
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("gmes.automail@changshininc.com", "GMES.AUTOMAIL", System.Text.Encoding.UTF8);
                mail.Bcc.Add("it.deny@changshininc.com");
                for (int i = 0; i < dtMail.Rows.Count; i++)
                {
                    if (dtMail.Rows[i]["TYPE"].ToString() == "TO")
                    {
                        mail.To.Add(dtMail.Rows[i]["EMAIL"].ToString());
                    }
                    else if (dtMail.Rows[i]["TYPE"].ToString() == "CC")
                    {
                        mail.CC.Add(dtMail.Rows[i]["EMAIL"].ToString());
                    }
                    else if (dtMail.Rows[i]["TYPE"].ToString() == "BCC")
                    {
                        mail.Bcc.Add(dtMail.Rows[i]["EMAIL"].ToString());
                    }
                }

                string htmlBody = fnGenerateHtml(dt);
                mail.Subject = "Official Car Request";
                mail.Body = htmlBody;
                mail.IsBodyHtml = true;
                mail.SubjectEncoding = System.Text.Encoding.UTF8;
                mail.BodyEncoding = System.Text.Encoding.UTF8;

                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);
                // Tambahkan gambar sebagai LinkedResource (inline)
                string destinationPath = Application.StartupPath.ToString() + "\\image.jpg";
                LinkedResource imageResource = new LinkedResource(destinationPath, MediaTypeNames.Image.Jpeg);
                imageResource.ContentId = "logoImage"; // Harus sama dengan yang di cid:
                imageResource.TransferEncoding = TransferEncoding.Base64;
                htmlView.LinkedResources.Add(imageResource);
                mail.AlternateViews.Add(htmlView);

                SmtpClient smtpServer = new SmtpClient("jjmail2.dskorea.com", 587);
                smtpServer.UseDefaultCredentials = false;
                smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpServer.Credentials = new System.Net.NetworkCredential("gmes.automail@dskorea.com", "csg1122!@");
                smtpServer.EnableSsl = true;
                System.Net.ServicePointManager.ServerCertificateValidationCallback += (s, cert, chain, sslPolicyErrors) => true;
                smtpServer.Send(mail);

                cnt++;
                this.Invoke((MethodInvoker)delegate
                {
                    lblSent.Text = cnt.ToString();
                });
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    lblError.Text = ex.Message;
                });
            }
        }

        private string fnGenerateHtml(DataTable dtData)
        {
            StringBuilder html = new StringBuilder();
            try
            {
                html.Append("<!DOCTYPE html>");
                html.Append("<html lang=\"en\">");
                html.Append("<head>");
                html.Append("<meta charset=\"UTF-8\">");
                html.Append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
                html.Append("<title>Car Arrangement Summary</title>");
                html.Append("<style>");
                html.Append("body { font-family: 'Times New Roman'; font-size: 18px;}");
                html.Append("h1 { font-family: 'Times New Roman', serif; font-style: italic; }");
                html.Append("table { width: 100%; border-collapse: collapse; margin-top: 20px; }");
                html.Append("th, td { border: 1px solid black; padding: 8px; text-align: center; font-family: 'Times New Roman'; font-size: 18px; font-style: italic;}");
                html.Append("th { border: 1px solid white; padding: 8px; text-align: center; }");
                html.Append("th { background-color: rgb(15, 0, 95); color: white; }");
                html.Append("img { display: block; margin-left: 0; margin-right: auto; max-width: 100%; height: auto; }");
                html.Append("</style>");
                html.Append("</head>");
                html.Append("<body>");

                //string imagePath = @"\\10.10.100.24\Gambar\image.jpg";
                destinationPath = Application.StartupPath.ToString() + "\\image.jpg";
                html.Append("<img src='cid:logoImage'>");
                //MessageBox.Show(destinationPath);

                // Tabel Jadwal Hari Ini
                html.Append("<h1>Informasi kendaraan yang dipesan akan dibagikan.</h1>");
                html.Append("<table>");
                html.Append("<tr><th>No</th><th>Departure Date</th><th>Time</th><th>Car No</th><th>Type</th><th>Driver</th><th>Driver's Phone No</th><th>Departure</th><th>Destination</th><th>Name</th></tr>");
                foreach (DataRow row in dtData.Rows)
                {
                    html.Append("<tr>");
                    foreach (var item in row.ItemArray)
                    {
                        html.AppendFormat("<td>{0}</td>", item);
                    }
                    html.Append("</tr>");
                }
                html.Append("</table>");
                html.Append("</body>");
                html.Append("</html>");
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    lblError.Text = ex.Message;
                });
            }
            return html.ToString();
        }

        private void tmrRun_Tick(object sender, EventArgs e)
        {
            string cek = DateTime.Now.ToString("HHmmss");
            string cekDay = DateTime.Now.ToString("ddd").ToUpper();

            if (cekDay.Contains("SAT") && cek == Properties.Settings.Default.SendTimeSat)
            {
                fnStart();
                System.Threading.Thread.Sleep(10000);
            }
            else if (!cekDay.Contains("SAT") && cek == Properties.Settings.Default.SendTime)
            {
                fnStart();
                System.Threading.Thread.Sleep(10000);
            }

            if (cek == "000000")
            {
                cnt = 0;
            }

        }

        private void fnStart()
        {
            try
            {
                if (start)
                {
                    if (!bgWork.IsBusy)
                    {
                        bgWork.RunWorkerAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    lblError.Text = ex.Message;
                });
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Select();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            allowClose = true;
            this.Close();
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Select();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!allowClose)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            }
        }
    }
}
