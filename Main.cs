using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DeliveryDate.Modules;
using System.Collections;
using System.Net.Mail;
using System.Net;
using System.IO;

namespace DeliveryDate
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private string sSQL, sEmail, sData, sLSR, sLSREmail;
        ArrayList agentsList = new ArrayList();
        private bool bSuccess;
        // TODO
        private static string stLogFile = @"d:\LogFiles\DelDate.txt";
        int iCounter, iSuccess;
        private static string stConn = "data source=MSSQL1;initial catalog=Reporting;password=15-uK*p;persist security info=True;user id=sa;Trusted_Connection=no;Application Name=DeliveryDateEmailer";
        private static string prodConn = "data source=NTREPORTS;initial catalog=lp01200;password=15-uK*p;persist security info=True;user id=sa;Trusted_Connection=no;Application Name=DeliveryDateEmailer";

        private void Main_Load(object sender, EventArgs e)
        {
            sSQL = "sp_DeliveryDate";
            cSQL helperSQL = new cSQL();
            DataSet fullDs = helperSQL.clsDataSet(sSQL, stConn);
            DataTable fullDt = fullDs.Tables[0];

            if (fullDt.Rows.Count > 0)
            {
                foreach (DataRow dr in fullDt.Rows)
                {
                    String thisAgent = dr["AGENT_NUMBER"].ToString().Trim();
                    if (agentsList.IndexOf(thisAgent) == -1)
                    {
                        DataRow[] agentRows = fullDt.Select("AGENT_NUMBER = '" + thisAgent + "'");

                        ArrayList policyList = new ArrayList();
                        String agentName;
                        String agentNFCode = agentRows[0]["SA_NAME_FORMAT"].ToString();                     

                        if (agentNFCode.CompareTo("B") == 0)
                        {
                            agentName = agentRows[0]["SA_BUSINESS"].ToString();
                        }
                        else
                        {
                            agentName = agentRows[0]["SA_FIRST"].ToString().Trim() + " " + agentRows[0]["SA_LAST"].ToString().Trim();
                        }

                        // header
                        sData = "<body bgcolor='#999999' style='margin:0;padding:0;'>";
                        sData = sData + "<div style='PADDING-RIGHT: 40px; PADDING-LEFT: 40px; PADDING-BOTTOM: 200px; WIDTH: 100%; PADDING-TOP: 40px; BACKGROUND-COLOR: #545454'>";
                        sData = sData + "<table cellspacing='0' cellpadding='0' width='200' align='center' border='0'>";
                        sData = sData + "<tbody>";
                        sData = sData + "<tr><td valign='top' colspan='3'><a href='http://www.emcnationallife.com/agent'><img alt='EMCNL' height='103' src='http://www.emcnationallife.com/contest/communiquecm/emailtemplateheadercm.jpg' width='640' border='0'/></a></td></tr>";
                        sData = sData + "<tr>";
                        sData = sData + "<td bgcolor='#ffffff' >&nbsp;</td>";
                        sData = sData + "<td bgcolor='#ffffff'>&nbsp;</td>";
                        sData = sData + "<td bgcolor='#ffffff'>&nbsp;</td></tr>";
                        sData = sData + "<tr>";
                        sData = sData + "<td bgcolor='#ffffff'>&nbsp;</td>";
                        sData = sData + "<td bgcolor='#ffffff'>&nbsp;</td>";
                        sData = sData + "<td bgcolor='#ffffff'>&nbsp;</td></tr>";
                        sData = sData + "<tr><td bgcolor='#ffffff'>&nbsp;</td><td bgcolor='#ffffff' colspan='2'>";
                        sData = sData + "<table style='font-family:Arial, Helvetica, sans-serif; font-size:10pt; width:100%;'><tr><td colspan='4'>" + agentName + " (" + agentRows[0]["AGENT_NUMBER"].ToString().Trim() + ")</td></tr>";
                        sData = sData + "<tr><td colspan='4'>&nbsp;</td></tr>";
                        sData = sData + "<tr><td colspan='4'>This email is being provided to advise you the following policy(s) have been mailed from EMC National Life. This information can also be accessed on our agent website at   <a href='https://www.EMCNationalLife.com'>www.EMCNationalLife.com</a></td></tr>";
                        sData = sData + "<tr><td colspan='4'>&nbsp;</td></tr>";
                        sData = sData + "<tr><td colspan='4'>&nbsp;</td></tr>";
                        sData = sData + "<tr style='text-decoration:underline'><td>OWNER</td><td>DESCRIPTION</td></tr>";
                        sData = sData + "<tr><td colspan='4'>&nbsp;</td></tr>";
                        for (int i = 0; i < agentRows.Length; i++)
                        {
                            String policy = agentRows[i]["POLICY_NUMBER"].ToString().Trim();
                            policyList.Add(policy);
                            String mailTo = agentRows[i]["MAIL_TO"].ToString().Trim();
                            if (string.IsNullOrEmpty(mailTo))
                            {
                                String product = agentRows[i]["PRODUCT_ID"].ToString().Trim();
                                String lobSQL = "";
                                lobSQL = "SELECT * FROM NTREPORTS.lp01200.dbo.PPRDF_EMCNL WHERE EMCNL_PRODUCT_ID = '" + product + "';";
                                cSQL clsLob = new cSQL();
                                DataSet dsLob = clsLob.clsDataSet(lobSQL, prodConn);
                                DataTable dtLob = dsLob.Tables[0];
                                if (dtLob.Rows.Count > 0) 
                                {
                                    DataRow drLob = dtLob.Rows[0];
                                    String prodType = drLob["EMCNL_INDIV_WORK"].ToString().Trim();
                                    if (prodType.Equals("I"))
                                    {
                                        mailTo = "AGENT";
                                    }
                                    else if (prodType.Equals("W")) 
                                    {
                                        mailTo = "POLICYOWNER";
                                    }
                                }
                            }
                            DateTime dtMailDate = DateTime.Parse(agentRows[i]["DELIVERY_DATE"].ToString());
                            String strMailDate = dtMailDate.ToShortDateString();
                            sData = sData + "<tr><td>" + agentRows[i]["OWNER_NAME"].ToString() + "</td><td>Policy Mailed To " + mailTo + "</td><td>" + strMailDate + "</td></tr>"; 
                        }
                        // now send the final email
                        // footer
                        sData = sData + "<tr><td colspan='4'>&nbsp;</td></tr>";
                        //sData = sData + "<tr><td colspan='4'>For complete details, please log in to our website at <a href='https://www.EMCNationalLife.com/AgentLogin.aspx?Page=Pend' target='_blank'>www.EMCNationalLife.com</a></td></tr>";
                        //sData = sData + "<tr><td colspan='4'>&nbsp;</td></tr>";
                        sData = sData + "<tr><td colspan='4'>If you have any questions, please contact us at 800-232-5818, Monday-Friday, 8 a.m.-4:30 p.m. (Central Time).  Thank you for your business.</td></tr>";
                        sData = sData + "<tr><td colspan='4'>&nbsp;</td></tr>";
                        sData = sData + "</table>";
                        // chart of abbreviations
                        //sData = sData + "<table style='font-family:Arial, Helvetica, sans-serif; font-size:8pt; width:100%;'><tr><td colspan='4' style='font-weight:bold;'>Requirements Reference Key</td></tr>";
                        //sData = sData + "<tr><td colspan='4'><a href='https://www.emcnationallife.com/PDFS/RequirementsReferenceKey.pdf' target='_blank'>Click here to view key</a></td></tr>";
                        //sData = sData + "</table>";
                        sData = sData + "</td></tr>";
                        sData = sData + "<tr>";
                        sData = sData + "<td width='40' bgcolor='#ffffff'>&nbsp;</td>";
                        sData = sData + "<td width='695' bgcolor='#ffffff'>&nbsp;</td>";
                        sData = sData + "<td width='40' bgcolor='#ffffff'>&nbsp;</td></tr>";
                        sData = sData + "<tr><td valign='top' colspan='3'><img alt='' height='30' src='http://www.emcnationallife.com/contest/communiquecm/emailtemplatefootercm.jpg' width='640'/></td></tr>";
                        sData = sData + "<tr>";
                        sData = sData + "<td colspan='3'>";
                        sData = sData + "<div align='center'>";
                        sData = sData + "<br/><p><font face='Arial, Helvetica, sans-serif' color='#ffffff' size='1'><strong>EMC National Life Company, 699 Walnut Street, Des Moines, IA 50309<br/>";
                        sData = sData + "This is an auto-generated e-mail. Please do not respond to this message for customer service issues.</strong></font></p>";
                        sData = sData + "</div></td></tr></tbody></table></div></body>";
                        iCounter = iCounter + 1;

                        agentsList.Add(thisAgent);
                        String emailOne = dr["EMAIL"].ToString().Trim(); // Email from table
                        String emailTwo = dr["EMAIL_TWO"].ToString().Trim(); // Email from Intranet
                        String sendTo;
                        if (string.Equals(emailOne, emailTwo, StringComparison.CurrentCultureIgnoreCase))
                        {
                            sendTo = emailTwo;
                        }
                        else if (string.IsNullOrEmpty(emailOne))
                        {
                            sendTo = emailTwo;
                        }
                        else if (string.IsNullOrEmpty(emailTwo))
                        {
                            sendTo = emailOne;
                        }
                        else
                        {
                            sendTo = emailTwo;
                        }
                   //     sendTo = "cadams@emcnl.com"; // For Test
                        sendTo = "JodiLarson@emcnl.com"; // For Test
                        bool bProceed = clsEmailer("communications@emcnl.com", sendTo, "Notice of Policy Delivery", "HTML", "Normal", sData, "", "", "");

                        if (bProceed)
                        {
                            iSuccess = iSuccess + 1;
                            String sqlUpd = "";
                            sqlUpd += "UPDATE MSSQL1.Reporting.dbo.tblPolDelivery SET EMAIL_SENT = 'Y' WHERE AGENT_NUMBER = '" + thisAgent + "' AND POLICY_NUMBER IN (";
                            for (int j = 0; j < policyList.Count; j++)
                            {
                                String thisPolicy = policyList[j].ToString();
                                if (policyList.Count - 1 == j)
                                {
                                    sqlUpd += "'" + thisPolicy + "');";
                                }
                                else
                                {
                                    sqlUpd += "'" + thisPolicy + "', ";
                                }
                            }
                            cSQL cUpd = new cSQL();
                            bool updSuccess = cUpd.clsUpdater(sqlUpd, stConn);
                        }
                    }
                }
                WriteLog(DateTime.Today.ToShortDateString() + " - Emails Generated:  " + iCounter.ToString() + ", " + iSuccess.ToString() + " successfully\r");
            }
            Application.ExitThread();
        }
        private bool clsEmailer(string sFrom, string sTo, string sSubject, string sFormat, string sPriority, string sData, string sBCC, string sCC, string sAttach)
        {
            string sErr;
            int iCounter = 0;
            // TODO SITE //
            //      sTo = "CMerritt@EMCNL.COM";
            sBCC = "";
            sCC = "";

            char crComma = Convert.ToChar(",");
            // clean input
            sFrom = sFrom.Replace(";", ",");
            sTo = sTo.Replace(";", ",");
            sBCC = sBCC.Replace(";", ",");
            sCC = sCC.Replace(";", ",");
            // create arrays
            string[] saCC = sCC.Split(crComma);
            int iCC = saCC.Length;
            string[] saBCC = sBCC.Split(crComma);
            int iBCC = saBCC.Length;

            MailMessage Mail = new MailMessage(sFrom, sTo, sSubject, sData);
            try
            {
                MailAddressCollection maColl = new MailAddressCollection();
                if (sCC != "")
                {
                    for (int x = 0; x < iCC; x++)
                    {
                        if (saCC[x].IndexOf("@") > 0)
                        {
                            maColl.Add(saCC[x]);
                            Mail.CC.Add(maColl[iCounter]);
                            iCounter = iCounter + 1;
                        }
                    }
                }
                if (sBCC != "")
                {
                    for (int y = 0; y < iBCC; y++)
                    {
                        if (saBCC[y].IndexOf("@") > 0)
                        {
                            maColl.Add(saBCC[y]);
                            Mail.Bcc.Add(maColl[iCounter]);
                            iCounter = iCounter + 1;
                        }
                    }
                }
                if (sPriority == "Normal")
                {
                    Mail.Priority = MailPriority.Normal;
                }
                else if (sPriority == "High")
                {
                    Mail.Priority = MailPriority.High;
                }
                else
                {
                    Mail.Priority = MailPriority.Low;
                }

                if (sFormat == "Text")
                {
                    Mail.IsBodyHtml = false;
                }
                else
                {
                    Mail.IsBodyHtml = true;
                }
                if (sAttach != "")
                {
                    Attachment myAttachment = new Attachment(sAttach);
                    Mail.Attachments.Add(myAttachment);
                    myAttachment = null;
                }
                ////// TODO SITE //     Test Here
                ////// need to put proper name here                
                /*      string server = "exchange";
                      SmtpClient client = new SmtpClient(server);
                      client.Send(Mail);
                      Mail.Dispose();
                      Mail = null;
                      return true; */

                // TODO SITE //
                // need to put proper name here
                /*   string server = "192.168.0.2";
                   SmtpClient client = new SmtpClient(server);
                   client.Port = 25;
                   client.UseDefaultCredentials = true;
                   client.Send(Mail);
                   Mail.Dispose();
                   Mail = null; 
                   return true; */

                SmtpClient client = new SmtpClient("smtp.office365.com", 25); // New Office 365 Sent from Server
                client.EnableSsl = true;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(sFrom, "N0w1sth3T1meFor");
                client.TargetName = "STARTTLS/smtp.office365.com";
                client.Send(Mail);
                return true;



            }
            catch (System.Exception e)
            {
                sErr = e.ToString();
                Mail.Dispose();
                Mail = null;
                return false;
            }
        }
        private void WriteLog(string sSummary)
        {
            StreamWriter file = new StreamWriter(stLogFile, true);
            file.WriteLine(sSummary);
            file.Close();
        }
    }
}
