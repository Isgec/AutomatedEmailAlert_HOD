using System;
using System.Data;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Net;
using System.Windows.Forms;
using System.Timers;
using ClosedXML;
using ClosedXML.Excel;

namespace AutomatedEmailAlert_HOD
{
    public partial class AutomatedEmailAlert_HOD : Form
    {
        public AutomatedEmailAlert_HOD()
        {
            InitializeComponent();
        }
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            this.button1_Click(null, null);
        }
        private static void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Environment.Exit(0);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
               
                using (SqlConnection con2 = new SqlConnection("Server=192.9.200.150; Database=IJTPerks; User Id=sa; Password=isgec12345;"))
                {
                    string sHODmail = "";
                    string sHODEmail = "";
                    string MailToHODs = "";
                    if (con2.State != ConnectionState.Open)
                    {
                        con2.Open();
                    }
                    SqlTransaction tran1 = con2.BeginTransaction();
                    string sDivisionCode = "select distinct Division from ProjectHODsAutoAlert";
                    SqlCommand cmdDivisionCode = new SqlCommand(sDivisionCode, con2, tran1);
                    var dt_Division = new System.Data.DataTable();
                    using (SqlDataReader dr_Division = cmdDivisionCode.ExecuteReader())
                    {
                        dt_Division.Load(dr_Division);
                        // return tb;
                    }
                    foreach (DataRow row_Division in dt_Division.Rows)
                    {
                        string DivisonCode = ("" + row_Division["Division"] + "").ToString();
                        string sHODCode = "select HODEmployeeCode from ProjectHODsAutoAlert where Division= '" + row_Division["Division"] + "'";
                        SqlCommand cmdHODCode = new SqlCommand(sHODCode, con2, tran1);
                        var dt3 = new System.Data.DataTable();
                        using (SqlDataReader dr2 = cmdHODCode.ExecuteReader())
                        {
                            dt3.Load(dr2);
                            // return tb;
                        }
                        if (dt3.Rows.Count > 1)
                        {
                            foreach (DataRow datarow1 in dt3.Rows)
                            {
                                sHODmail = "select EMailID from HRM_Employees where CardNo = " + datarow1["HODEmployeeCode"] + "";
                                SqlCommand cmdHODEmail = new SqlCommand(sHODmail, con2, tran1);
                                sHODEmail += cmdHODEmail.ExecuteScalar().ToString();
                                sHODEmail += ";";
                            }
                        }
                        else
                        {
                            sHODEmail = ";";
                        }
                        MailToHODs = sHODEmail;
                        //"pankaj.gupta@isgec.co.in;sagar.shukla@isgec.co.in";
                        //"hariharan.m@isgec.co.in";
                        //sHODEmail;
                        //"sagar.shukla@isgec.co.in";
                        // sHODEmail = "pankaj.gupta@isgec.co.in";
                        //sBuyerEmail + ";" + sIndenterEmail;
                        using (SqlConnection con = new SqlConnection("Server=192.9.200.129; Database=inforerpdb; User Id=dev1; Password=Dev1@12345;"))
                        {
                          
                          
                            var dt = new System.Data.DataTable();
                            string sToday = DateTime.Now.ToString("yyyy-MM-dd");
                            string sReportStartDate = DateTime.Now.AddDays(-15).ToString("yyyy-MM-dd");
                            string sRecord = @"select  distinct dpur401.t_orno as PONo,dpur201.t_cprj as Project, ltrim(rtrim(ttc052.t_dsca))+' '+ltrim(rtrim(ttc052.t_dscb)) 
                                     as ProjectName,dpur401.t_pono as POLine,
                            dpur401.t_otbp as SupplierCode,tcc100.t_nama as SupplierName,
                            Convert(varchar(20), dpur401.t_ddta,105) as PlannedDeliveryDate,
                            dpur201.t_rqno as IndentNo,dpur201.t_pono as IndentLineNo,
                            dpur401.t_item as ItemCode, cibd001.t_dsca as ItemDesc,
                            Convert(varchar(20), dpur201.t_dldt,105) as IndentPlannedDeliveryDate,
                            dpur400.t_ccon as Buyer,dpur200.t_remn as Indenter,dpur200.t_rdep as Division,
                            DATEDIFF(day,dpur201.t_dldt,dpur401.t_ddta) as DayDiff,
                            dmsl40.t_updt as UpdateDate
                            from ttdmsl400200 dmsl40
                             inner join ttdpur400200 dpur400 on dmsl40.t_orno=dpur400.t_orno
                            inner join ttdpur402200 dpur402 on dpur400.t_orno= dpur402.t_orno 
                            inner join ttdpur401200 dpur401 on dpur402.t_orno = dpur401.t_orno
                            inner join ttcibd001200  cibd001 on cibd001.t_item=dpur401.t_item
                            inner join ttdpur201200  dpur201 on dpur201.t_item=dpur401.t_item
                            inner join ttdpur202200 dpur202 on dpur202.t_rqno=dpur201.t_rqno
                            inner join ttdpur200200 dpur200 on dpur200.t_rqno= dpur201.t_rqno
                            inner join ttccom100200 tcc100 on tcc100.t_bpid =dpur401.t_otbp
                            inner join ttcmcs052200 ttc052 on ttc052.t_cprj=dpur201.t_cprj
                            where DATEDIFF(day,dpur201.t_dldt,dpur401.t_ddta)>0
                            and cibd001.t_kitm=1 and  CONVERT(date, dmsl40.t_updt) >= '" + sReportStartDate + @"'
                            and CONVERT(date, dmsl40.t_updt) <='" + sToday + @"' and dpur201.t_rqno= dpur402.t_rqno
                            and dpur402.t_pono= dpur401.t_pono and dpur201.t_item=cibd001.t_item
                            and dpur202.t_ppon=dpur401.t_pono and dpur201.t_pono= dpur202.t_pono
							and dpur200.t_rdep='" + row_Division["Division"] + @"'";

                            if (con.State != ConnectionState.Open)
                            {
                                con.Open();
                            }
                            SqlTransaction tran = con.BeginTransaction();
                            SqlCommand cmd = new SqlCommand(sRecord, con, tran);
                            using (SqlDataReader dr = cmd.ExecuteReader())
                            {
                                dt.Load(dr);
                            }
                            dt.TableName = "HOD_ProjDeviation";
                            if (dt.Rows.Count > 0)
                            {
                                ExportDataSetToExcel(dt, DivisonCode);

                                SendMail(MailToHODs, DivisonCode);
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                System.Timers.Timer timer = new System.Timers.Timer(5000);
                timer.Elapsed += Timer_Elapsed;
                timer.Start();
            }
            finally
            {
                System.Timers.Timer timer = new System.Timers.Timer(5000);
                timer.Elapsed += Timer_Elapsed;
                timer.Start();
            }
        }

        private void ExportDataSetToExcel(System.Data.DataTable dt, string Division)
        {
            string AppLocation = "E:";
            AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            AppLocation = AppLocation.Replace("file:\\", "");
            string file = AppLocation + "\\ExcelFiles\\ProjectDeviation("+Division+"_" + DateTime.Now.AddDays(-15).ToString("dd-MM-yy") + " to " + DateTime.Now.ToString("dd-MM-yy") + ").xlsx";
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                wb.SaveAs(file);
            }
        }

        public void SendMail(string sHODEmail, string Division)
        {
            try
            {
                MailMessage mM = new MailMessage();
                mM.From = new MailAddress("baansupport@isgec.co.in");
                //foreach (var address in MailToBuyer_Indenter.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                //{
                //    mM.To.Add(address);
                //}
                foreach (var address in sHODEmail.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mM.To.Add(address);
                }
                // Below two Cc Added to check the proper email functionality
                //  mM.CC.Add("baansupport@isgec.co.in");
                // mM.CC.Add("veena@isgec.co.in");
                mM.Subject = "Deviation in Plan Delivery Date ";
                string AppLocation = "E:";
                AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                AppLocation = AppLocation.Replace("file:\\", "");
                string file = AppLocation + "\\ExcelFiles\\ProjectDeviation("+Division+"_" + DateTime.Now.AddDays(-15).ToString("dd-MM-yy") + " to " + DateTime.Now.ToString("dd-MM-yy") + ").xlsx";
                //string Project = "";


                // mM.Subject = "Deviation in Plan Delivery Date";

                mM.IsBodyHtml = true;
                mM.Body += "Please Find the attached document for PO Plan Delivery Deviation sheet.";
                mM.Body += "<br /><br /><br /><br /><br />Note- This is an autogenerated e-mail";
                mM.Body = mM.Body.Replace("\n", "<br />");
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(file); //Attaching File to Mail  
                mM.Attachments.Add(attachment);
                SmtpClient sC = new SmtpClient("192.9.200.214", 25);
                sC.DeliveryMethod = SmtpDeliveryMethod.Network;
                sC.UseDefaultCredentials = false;
                sC.Credentials = new NetworkCredential("baansupport@isgec.co.in", "isgec");
                sC.EnableSsl = false;  // true
                sC.Timeout = 10000000;
                sC.Send(mM);

            }
            catch (Exception ex)
            {
            }
        }
    }
}
