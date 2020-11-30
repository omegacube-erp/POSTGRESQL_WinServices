using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Configuration;
using System.Configuration.Internal;
using Microsoft.ServiceModel.Channels.Mail.ExchangeWebService.Exchange2007;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Threading;
using System.Runtime.InteropServices;
using FE_SymmetricNamespace;

namespace Edge_EmailService_pg
{
    public partial class Edge_EmailService : ServiceBase
    {
        public string strFromName = string.Empty;
        public string strSmtpServer = string.Empty;
        public string strSmtpPort = string.Empty;
        public string strSmtpUid = string.Empty;
        public string strSmtpPwd = string.Empty;
        public string strFromEmailID = string.Empty;
        public string strDomain = string.Empty;
        public string strExorSm = string.Empty;
        public string strOracleRun = string.Empty;
        public string strOracleReportPath = string.Empty;

        public string strEnableSSL = string.Empty;
        public string strUseDomainInSenderEmailId = string.Empty;

        public string strPreserveLog = string.Empty;

        
        public int nSchemas = 0;

        public string fName = string.Empty;

        protected static string AppPath = System.AppDomain.CurrentDomain.BaseDirectory;
        protected static string strLogFilePath = AppPath + "\\" + "Edge_EmailLog.txt";
        private static StreamWriter sw = null;

        public string selectQuery = string.Empty;

        public static int LoopCount;


        public Edge_EmailService()
        {
            InitializeComponent();
            
            nSchemas = Convert.ToInt16(ConfigurationManager.AppSettings["NoOfSchemas"]);

            strFromName = ConfigurationManager.AppSettings["frmName"];
            strFromEmailID = ConfigurationManager.AppSettings["frmEmailID"];

            strSmtpServer = ConfigurationManager.AppSettings["smtpServer"];
            strSmtpPort = ConfigurationManager.AppSettings["smtpPort"];

            strSmtpUid = ConfigurationManager.AppSettings["smtpUid"];
            strSmtpPwd = ConfigurationManager.AppSettings["smtpPwd"];

            strDomain = ConfigurationManager.AppSettings["Domain"];

            strExorSm = ConfigurationManager.AppSettings["ExhangeOrSMTP"];

            strOracleRun = ConfigurationManager.AppSettings["oracleRun"];
            strOracleReportPath = ConfigurationManager.AppSettings["oracleReportsPath"];

            strEnableSSL = ConfigurationManager.AppSettings["EnableSSL"];
            strUseDomainInSenderEmailId = ConfigurationManager.AppSettings["UseDomainInSenderEmailId"];

            strPreserveLog = ConfigurationManager.AppSettings["PreserveLog"];
             
        }

        OleDbDataAdapter da;
        OleDbConnection dbcon;
        OleDbDataAdapter dbad = new OleDbDataAdapter();
        

        public string strSno = string.Empty;
        public string strClient = string.Empty;
        public string strToAdd = string.Empty;
        public string strCcAdd = string.Empty;
        public string strBccAdd = string.Empty;
        public string strFromAdd = string.Empty;
        public string strMailSub = string.Empty;
        public string strMailBody = string.Empty;
        public string strAttachments = string.Empty;
        public string strNewsLetter = string.Empty;

        public string strReportName = string.Empty;
        public string strViewName = string.Empty;
        public string strCondParam = string.Empty;
        public string strCondValue = string.Empty;
        public string strReportFileName = string.Empty;

        string strUpdate = string.Empty;

        System.Timers.Timer testTimer = new System.Timers.Timer();

        OleDbCommand cmd = new OleDbCommand();


        protected override void OnStart(string[] args)
        {
            try
            {
                //testTimer.Enabled = false;
                //SendBulkEmail();

                if (string.IsNullOrEmpty(strUseDomainInSenderEmailId))
                {
                    strUseDomainInSenderEmailId = "N";
                }

                if (string.IsNullOrEmpty(strEnableSSL))
                {
                    strEnableSSL = "N";
                }

                if (string.IsNullOrEmpty(strPreserveLog))
                {
                    strPreserveLog = "N";
                }

                testTimer.Interval = 10000;
                testTimer.Elapsed += new System.Timers.ElapsedEventHandler(testTimer_Elapsed);
                testTimer.Enabled = true;
                LogExceptions(strLogFilePath, null, "Service Started");
            }
            catch (Exception ex)
            {

                LogExceptions(strLogFilePath, ex, null);
            }
        }

        protected override void OnStop()
        {
            try
            {
                Thread.Sleep(1000);
                testTimer.Enabled = false;
                LogExceptions(strLogFilePath, null, "Service Stopped");
            }
            catch (Exception ex)
            {

                LogExceptions(strLogFilePath, ex, null);
            }
        }


        private void testTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                LoopCount = LoopCount + 1;
                if (LoopCount == 3)
                {
                    LoopCount = 0;
                    testTimer.Enabled = false;
                    SendBulkEmailForMultiSchemas();
                    testTimer.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                LogExceptions(strLogFilePath, ex, null);
            }

        }

        public void SendBulkEmailForMultiSchemas()
        {
            try
            {
                if (nSchemas == 0)
                {
                    nSchemas = 10;
                }

                for (int i = 1; i <= nSchemas; i++)
                {
                    string strServerName = string.Empty;
                    string strUserName = string.Empty;
                    string strPassWord = string.Empty;

                    strServerName = ConfigurationManager.AppSettings["Server" + i];
                    strUserName = ConfigurationManager.AppSettings["User" + i];
                    strPassWord = ConfigurationManager.AppSettings["Pwd" + i];

                    if (!string.IsNullOrEmpty(strServerName))
                    {
                        if (!string.IsNullOrEmpty(strUserName))
                        {
                            if (!string.IsNullOrEmpty(strPassWord))
                            {
                                try
                                {
                                    LogExceptions(strLogFilePath, null, "Start Sending for Schema :- " + strUserName);
                                    SendBulkEmail(strServerName, strUserName, strPassWord);
                                    LogExceptions(strLogFilePath, null, "End Sending for Schema :- " + strUserName);
                                    Thread.Sleep(1000);
                                }
                                catch(Exception exs)
                                {
                                    Insert_Sys_Log(exs.Message.ToString(), strServerName, strUserName, strPassWord);
                                    LogExceptions(strLogFilePath, exs, null);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogExceptions(strLogFilePath, ex, null);
            }
        }

        public void SendBulkEmail(string serName, string uName, string pWord)
        {

            try
            {
                DataSet dsNew1 = new DataSet();
                selectQuery = string.Empty;
                selectQuery = "SELECT SNO FROM SYS_GENERIC_EMAIL WHERE SEND_CONFIRMATION='N'";
                dbcon = new OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + serName + ";User Id=" + uName + ";Password=" + pWord + ";");
                da = new OleDbDataAdapter(selectQuery, dbcon);

                da.Fill(dsNew1, "SNO");

                if (dsNew1.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < dsNew1.Tables[0].Rows.Count; i++)
                    {

                        DataSet dsNew = new DataSet();

                        selectQuery = string.Empty;
                        selectQuery = "SELECT SNO,TO_ADDRESS,CC_ADDRESS,BCC_ADDRESS,MAIL_SUBJECT,ATTACHMENT,REPORT_NAME,VIEW_NAME,CONDITION_PARAMETER,CONDITION_VALUE,SEND_CONFIRMATION,CREATED_BY,CREATED_DATE,CHANGED_BY,CHANGED_DATE,NEWS_LETTER,REPORT_FILE_NAME,FROM_ID,MAIL_BODY,SOURCE_NAME,SOURCE_TYPE FROM SYS_GENERIC_EMAIL WHERE SNO ='" + dsNew1.Tables[0].Rows[i]["SNO"].ToString() + "'";
                        //string selectQuery = "SELECT * FROM SYS_GENERIC_EMAIL WHERE SEND_CONFIRMATION='N'";
                        //dbcon = new OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + serName + ";User Id=" + uName + ";Password=" + pWord + ";");
                        da = new OleDbDataAdapter(selectQuery, dbcon);

                        da.Fill(dsNew, "Email");

                        //for (int i = 0; i < dsNew.Tables[0].Rows.Count; i++)
                        //{
                        bool chkEmail = false;

                        strSno = string.Empty;
                        strClient = string.Empty;

                        strToAdd = string.Empty;
                        strCcAdd = string.Empty;
                        strBccAdd = string.Empty;
                        strFromAdd = string.Empty;
                        strMailSub = string.Empty;
                        strMailBody = string.Empty;
                        strAttachments = string.Empty;
                        strNewsLetter = string.Empty;

                        strReportName = string.Empty;
                        strCondParam = string.Empty;
                        strCondValue = string.Empty;
                        strReportFileName = string.Empty;
                        strViewName = string.Empty;

                        strSno = dsNew.Tables[0].Rows[i]["SNO"].ToString();
                        //strClient = dsNew.Tables[0].Rows[i]["CLIENT"].ToString();

                        strToAdd = dsNew.Tables[0].Rows[i]["TO_ADDRESS"].ToString();
                        strCcAdd = dsNew.Tables[0].Rows[i]["CC_ADDRESS"].ToString();
                        strBccAdd = dsNew.Tables[0].Rows[i]["BCC_ADDRESS"].ToString();
                        strFromAdd = dsNew.Tables[0].Rows[i]["FROM_ID"].ToString();
                        strMailSub = dsNew.Tables[0].Rows[i]["MAIL_SUBJECT"].ToString();
                        strMailBody = dsNew.Tables[0].Rows[i]["MAIL_BODY"].ToString();
                        strAttachments = dsNew.Tables[0].Rows[i]["ATTACHMENT"].ToString();
                        strNewsLetter = dsNew.Tables[0].Rows[i]["NEWS_LETTER"].ToString();

                        strReportName = dsNew.Tables[0].Rows[i]["REPORT_NAME"].ToString();
                        strCondParam = dsNew.Tables[0].Rows[i]["CONDITION_PARAMETER"].ToString();
                        strCondValue = dsNew.Tables[0].Rows[i]["CONDITION_VALUE"].ToString();
                        strReportFileName = dsNew.Tables[0].Rows[i]["REPORT_FILE_NAME"].ToString();
                        strViewName = dsNew.Tables[0].Rows[i]["VIEW_NAME"].ToString();


                        //if (string.IsNullOrEmpty(strFromAdd))
                        //{
                        //    strFromAdd = string.Empty;
                        //    strFromAdd = strFromEmailID;
                        //}

                        //if (strFromAdd==string.Empty)
                        //{
                        //    strFromAdd = strFromEmailID;
                        //}

                        if (string.IsNullOrEmpty(strFromAdd))
                        {
                            strFromAdd = string.Empty;
                            strFromAdd = strFromEmailID;
                        }


                        //chkEmail = SendEmail(dsNew.Tables[0].Rows[i]["TO_ADDRESS"].ToString(),dsNew.Tables[0].Rows[i]["CC_ADDRESS"].ToString(),dsNew.Tables[0].Rows[i]["BCC_ADDRESS"].ToString(),
                        if (!string.IsNullOrEmpty(strToAdd))
                        {
                            if (strExorSm == "E")
                            {
                                // Using Exchange
                                chkEmail = SendEmailWithExchange(strToAdd, strCcAdd, strBccAdd, strFromAdd, strMailSub, strMailBody, strAttachments, serName, uName, pWord);
                            }
                            else
                            {
                                // Using  SMTP
                                chkEmail = SendEmailWithSMTP(strToAdd, strCcAdd, strBccAdd, strFromAdd, strMailSub, strMailBody, strAttachments, serName, uName, pWord);
                            }

                            if (chkEmail == true)
                            {
                                strUpdate = "UPDATE SYS_GENERIC_EMAIL  SET SEND_CONFIRMATION='Y',CHANGED_BY='ADMN', CHANGED_DATE=SYSDATE WHERE SNO =" + strSno;
                                LogExceptions(strLogFilePath, null, "Mail has been Send to:" + strSno);
                            }
                            else
                            {
                                strUpdate = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='R',CHANGED_BY='ADMN', CHANGED_DATE=SYSDATE WHERE SNO =" + strSno;
                            }
                        }
                        else
                        {
                            strUpdate = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='R',CHANGED_BY='ADMN', CHANGED_DATE=SYSDATE  WHERE SNO=" + strSno;
                        }
                        cmd.Connection = dbcon;
                        cmd.CommandText = strUpdate;
                        cmd.CommandType = CommandType.Text;
                        cmd.Connection.Open();
                        int count = cmd.ExecuteNonQuery();
                        cmd.Connection.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Insert_Sys_Log(ex.Message.ToString(), serName, uName, pWord);
                LogExceptions(strLogFilePath, ex, null);
            }
        }


        public void LogExceptions(string filePath, [Optional, DefaultParameterValue(null)]Exception ex, [Optional, DefaultParameterValue(null)]string msg)
        {
            if (File.Exists(filePath))
            {
                if (strPreserveLog == "N")
                {
                    File.Delete(filePath);
                }
            }
            if (false == File.Exists(filePath))
            {
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                fs.Close();
            }
            WriteExceptionLog(filePath, ex, msg);
        }

        private static void WriteExceptionLog(string strPathName, [Optional, DefaultParameterValue(null)]Exception objException, [Optional, DefaultParameterValue(null)]string msg)
        {
            if (string.IsNullOrEmpty(msg))
            {
                sw = new StreamWriter(strPathName, true);
                sw.WriteLine("^^-------------------------------------------------------------------^^");
                sw.WriteLine("Source		: " + objException.Source.ToString().Trim());
                sw.WriteLine("Method		: " + objException.TargetSite.Name.ToString());
                sw.WriteLine("Date		: " + DateTime.Now.ToLongTimeString());
                sw.WriteLine("Time		: " + DateTime.Now.ToShortDateString());
                sw.WriteLine("Error		: " + objException.Message.ToString().Trim());
                sw.WriteLine("Stack Trace	: " + objException.StackTrace.ToString().Trim());
                sw.WriteLine("^^-------------------------------------------------------------------^^");
            }
            else
            {
                sw = new StreamWriter(strPathName, true);
                sw.WriteLine(msg);
            }
            sw.Flush();
            sw.Close();
        }

        public void Cert()
        {
            ServicePointManager.ServerCertificateValidationCallback =
                  delegate(Object obj, X509Certificate certificate,
                  X509Chain chain, SslPolicyErrors errors)
                  {
                      // Replace this line with code to validate server    
                      // certificate.   
                      return true;
                  };
        }

        public bool SendEmailWithExchange(string strTo, string strCc, string strBcc, string strFrom, string strSubject, string strBody, string strAttachmentPath, string sName, string sUser, string sPwd)
        {
            //char[] splitter1 = { ',' };
            char[] splitter = new char[] { ',', ';', ':' };
            ExchangeServiceBinding ewsServiceBinding = new ExchangeServiceBinding();

            string[] SmtpUserId = strSmtpUid.Split('@');
            string SmtpEmailId = string.Empty;
            string SmtpPassword1 = string.Empty;
            string SmtpDomain = string.Empty;

            if (SmtpUserId.Length == 2)
            {
                SmtpEmailId = SmtpUserId[0];
                SmtpDomain = SmtpUserId[1];
            }

            string SmtpEmailValue = string.Empty;


            //if (strUseDomainInSenderEmailId == "Y")
            //{
            //    SmtpEmailValue = strSmtpUid;
            //}
            //else
            //{
            //    SmtpEmailValue = SmtpEmailId;
            //}

            if (strUseDomainInSenderEmailId == "Y")
            {
                SmtpEmailValue = strFrom;
            }
            else
            {
                SmtpEmailValue = strSmtpUid;
            }

            if (!string.IsNullOrEmpty(strFrom))
            {
                DataSet dsUser = new DataSet();

                string selectQuery = "SELECT E_MAIL,EMAIL_PASSWORD FROM SYS_USER_MASTER WHERE UPPER(E_MAIL)=upper('" + strFrom.ToUpper().Trim() + "')";
                LogExceptions(strLogFilePath, null, selectQuery);
                dbcon = new OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + sName + ";User Id=" + sUser + ";Password=" + sPwd + ";");
                da = new OleDbDataAdapter(selectQuery, dbcon);
                dsUser.Clear();
                da.Fill(dsUser, "Email");

                string[] EmailElements = strFrom.Split('@');
                string EmailId = string.Empty;
                string Password1 = string.Empty;
                string Domain = string.Empty;
                if (EmailElements.Length == 2)
                {
                    EmailId = EmailElements[0];
                    Domain = EmailElements[1];
                }

                if (!string.IsNullOrEmpty(EmailId) && !string.IsNullOrEmpty(Domain))
                {
                    if (dsUser != null)
                    {
                        if (dsUser.Tables[0].Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dsUser.Tables[0].Rows[0]["E_MAIL"].ToString()))
                            {
                                if (!string.IsNullOrEmpty(dsUser.Tables[0].Rows[0]["EMAIL_PASSWORD"].ToString()))
                                {
                                    Password1 = Decrypt(dsUser.Tables[0].Rows[0]["EMAIL_PASSWORD"].ToString());
                                    if (strUseDomainInSenderEmailId == "Y")
                                    {
                                        ewsServiceBinding.Credentials = new NetworkCredential(EmailId, Password1, strDomain);
                                    }
                                    else
                                    {
                                        ewsServiceBinding.Credentials = new NetworkCredential(EmailId, Password1, strDomain);
                                    }
                                }
                                else
                                {
                                    //ewsServiceBinding.Credentials = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                                    ewsServiceBinding.Credentials = new NetworkCredential(SmtpEmailValue, strSmtpPwd, strDomain);
                                }
                            }
                            else
                            {
                                //ewsServiceBinding.Credentials = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                                ewsServiceBinding.Credentials = new NetworkCredential(SmtpEmailValue, strSmtpPwd, strDomain);
                            }
                        }
                        else
                        {
                            //ewsServiceBinding.Credentials = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                            ewsServiceBinding.Credentials = new NetworkCredential(SmtpEmailValue, strSmtpPwd, strDomain);
                        }
                    }
                    else
                    {
                        //ewsServiceBinding.Credentials = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                        ewsServiceBinding.Credentials = new NetworkCredential(SmtpEmailValue, strSmtpPwd, strDomain);
                    }
                }
                else
                {
                    //ewsServiceBinding.Credentials = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                    ewsServiceBinding.Credentials = new NetworkCredential(SmtpEmailValue, strSmtpPwd, strDomain);
                }
            }
            else
            {
                //ewsServiceBinding.Credentials = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                ewsServiceBinding.Credentials = new NetworkCredential(SmtpEmailValue, strSmtpPwd, strDomain);
            }

            //ewsServiceBinding.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
            ewsServiceBinding.Url = @"https://" + strSmtpServer + "/EWS/exchange.asmx";
            Cert();
            MessageType emMessage = new MessageType();

            //emMessage.From = new SingleRecipientType();
            //emMessage.From.Item = new EmailAddressType();
            //emMessage.From.Item.EmailAddress = strFrom;

            List<EmailAddressType> recipients = new List<EmailAddressType>();

            // For Reply Field
            Array arrReply;
            arrReply = strFrom.Split(splitter);
            recipients.Clear();
            foreach (string s in arrReply)
            {
                EmailAddressType recipient = new EmailAddressType();
                recipient.EmailAddress = s;
                recipients.Add(recipient);
            }

            emMessage.ReplyTo = recipients.ToArray();


            // For To Field
            Array arrTo;
            arrTo = strTo.Split(splitter);
            recipients.Clear();
            foreach (string s in arrTo)
            {
                EmailAddressType recipient = new EmailAddressType();
                recipient.EmailAddress = s;
                recipients.Add(recipient);
            }

            emMessage.ToRecipients = recipients.ToArray();

            // For CC              
            if (!string.IsNullOrEmpty(strCc))
            {
                Array arrCc;
                arrCc = strCc.Split(splitter);

                recipients = new List<EmailAddressType>();
                recipients.Clear();
                foreach (string s in arrCc)
                {
                    EmailAddressType recipient = new EmailAddressType();
                    recipient.EmailAddress = s;
                    recipients.Add(recipient);
                }
                emMessage.CcRecipients = recipients.ToArray();
            }

            // For Bcc              
            if (!string.IsNullOrEmpty(strBcc))
            {
                Array arrBcc;
                arrBcc = strBcc.Split(splitter);

                recipients = new List<EmailAddressType>();
                recipients.Clear();
                foreach (string s in arrBcc)
                {
                    EmailAddressType recipient = new EmailAddressType();
                    recipient.EmailAddress = s;
                    recipients.Add(recipient);
                }
                emMessage.BccRecipients = recipients.ToArray();
            }


            emMessage.Subject = strSubject;
            emMessage.Body = new BodyType();
            emMessage.Body.BodyType1 = BodyTypeType.HTML;
            emMessage.Body.Value = strBody;

            //emMessage.Body.Value = strBody;
            //if (!string.IsNullOrEmpty(strAttachmentPath))
            //{
            //    emMessage.Body.Value = strBody + "<BR>" + ReadHtmlFile(strAttachmentPath) + "<BR>" + UnSubscribe(strTo);
            //}
            //else
            //{
            //    emMessage.Body.Value = strBody;
            //}

            emMessage.ItemClass = "IPM.Note";
            emMessage.Sensitivity = SensitivityChoicesType.Normal;

            try
            {
                ItemIdType iiCreateItemid = CreateDraftMessage(ewsServiceBinding, emMessage);

                // For Attachments
                if (!string.IsNullOrEmpty(strAttachmentPath))
                {
                    Array arrAttach;
                    arrAttach = strAttachmentPath.Split(splitter[0]);

                    foreach (string s in arrAttach)
                    {
                        if (File.Exists(s))
                        {
                            iiCreateItemid = CreateAttachment(ewsServiceBinding, s, iiCreateItemid);
                        }
                    }
                }

                // For NewsLetterAttachment
                if (strNewsLetter == "R")
                {
                    string FileName = string.Empty;
                    FileName = NewsLetterAttachment(sName, sUser, sPwd);

                    if (!string.IsNullOrEmpty(FileName))
                    {
                        if (File.Exists(FileName))
                        {
                            iiCreateItemid = CreateAttachment(ewsServiceBinding, FileName, iiCreateItemid);
                        }
                    }
                }

                SendMessage(ewsServiceBinding, iiCreateItemid);
                return true;
            }
            catch (Exception ex)
            {
                Insert_Sys_Log(ex.Message.ToString(), sName, sUser, sPwd);
                LogExceptions(strLogFilePath, ex, null);
                return false;
            }
        }

        public bool SendEmailWithSMTP(string strTo, string strCc, string strBcc, string strFrom, string strSubject, string strBody, string strAttachmentPath, string sName, string sUser, string sPwd)
        {
            try
            {
                //char[] splitter = { ',' };
                char[] splitter = new char[] { ',', ';', ':' };

                MailMessage mm = new MailMessage();

                //mm.From = new MailAddress(strSmtpUid, strFromName);
                //strFromEmailID
                if (!string.IsNullOrEmpty(strFrom.Trim()))
                {
                    mm.ReplyTo = new MailAddress(strFrom);
                }
                else
                {
                    mm.ReplyTo = new MailAddress(strFromEmailID);
                }

                // For To Field
                Array arrTo;
                arrTo = strTo.Split(splitter);
                foreach (string s in arrTo)
                {
                    mm.To.Add(s);
                }

                // For Cc Field
                if (!string.IsNullOrEmpty(strCc))
                {
                    Array arrCc;
                    arrCc = strCc.Split(splitter);
                    foreach (string s in arrCc)
                    {
                        mm.CC.Add(s);
                    }
                }

                // For Bcc Field
                if (!string.IsNullOrEmpty(strBcc))
                {
                    Array arrBcc;
                    arrBcc = strBcc.Split(splitter);
                    foreach (string s in arrBcc)
                    {
                        mm.Bcc.Add(s);
                    }
                }
                // For Subject
                mm.Subject = strSubject;

                // For Attachments
                if (!string.IsNullOrEmpty(strAttachmentPath))
                {
                    Array arrToAttach;
                    arrToAttach = strAttachmentPath.Split(splitter[0]);

                    if (arrToAttach.Length != 0)
                    {
                        foreach (string s1 in arrToAttach)
                        {
                            if (File.Exists(s1))
                            {
                                Attachment attachFile = new Attachment(s1);
                                mm.Attachments.Add(attachFile);
                            }
                        }
                    }
                }

                // For Body
                mm.IsBodyHtml = true;
                mm.Body = strBody;

                // For News Letters            
                if (strNewsLetter == "R")
                {
                    string FileName = string.Empty;
                    FileName = NewsLetterAttachment(sName, sUser, sPwd);

                    if (!string.IsNullOrEmpty(FileName))
                    {
                        if (File.Exists(FileName))
                        {
                            mm.Attachments.Add(new Attachment(FileName));
                        }
                    }
                }

                //mm.Body = strBody;

                SmtpClient smtpC = new SmtpClient();
                NetworkCredential netCred;

                string[] SmtpUserId = strSmtpUid.Split('@');
                string SmtpEmailId = string.Empty;
                string SmtpPassword1 = string.Empty;
                string SmtpDomain = string.Empty;

                if (SmtpUserId.Length == 2)
                {
                    SmtpEmailId = SmtpUserId[0];
                    SmtpDomain = SmtpUserId[1];
                }

                string SmtpEmailValue = string.Empty;


                //if (strUseDomainInSenderEmailId == "Y")
                //{
                //    SmtpEmailValue = strSmtpUid;
                //}
                //else
                //{
                //    SmtpEmailValue = SmtpEmailId;
                //}

                if (strUseDomainInSenderEmailId == "Y")
                {
                    SmtpEmailValue = strFrom;
                }
                else
                {
                    SmtpEmailValue = strSmtpUid;
                }


                //NetworkCredential netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                if (!string.IsNullOrEmpty(strFrom))
                {
                    DataSet dsUser = new DataSet();

                    string selectQuery = "SELECT E_MAIL,EMAIL_PASSWORD,FIRST_NAME || LAST_NAME NAME  FROM SYS_USER_MASTER WHERE upper(E_MAIL)=upper('" + strFrom.Trim() + "')";
                    LogExceptions(strLogFilePath, null, selectQuery);
                    dbcon = new OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + sName + ";User Id=" + sUser + ";Password=" + sPwd + ";");
                    da = new OleDbDataAdapter(selectQuery, dbcon);
                    dsUser.Clear();
                    da.Fill(dsUser, "Email");

                    string[] EmailElements = strFrom.Split('@');
                    string EmailId = string.Empty;
                    string Password1 = string.Empty;
                    string Domain = string.Empty;
                    if (EmailElements.Length == 2)
                    {
                        EmailId = EmailElements[0];
                        Domain = EmailElements[1];
                    }

                    if (!string.IsNullOrEmpty(EmailId) && !string.IsNullOrEmpty(Domain))
                    {
                        if (dsUser != null)
                        {
                            if (dsUser.Tables[0].Rows.Count > 0)
                            {
                                if (!string.IsNullOrEmpty(dsUser.Tables[0].Rows[0]["E_MAIL"].ToString()))
                                {
                                    if (!string.IsNullOrEmpty(dsUser.Tables[0].Rows[0]["EMAIL_PASSWORD"].ToString()))
                                    {
                                        Password1 = Decrypt(dsUser.Tables[0].Rows[0]["EMAIL_PASSWORD"].ToString());
                                        mm.From = new MailAddress(strFrom, dsUser.Tables[0].Rows[0]["NAME"].ToString());
                                        if (strUseDomainInSenderEmailId == "Y")
                                        {
                                            netCred = new NetworkCredential(strFrom, Password1, strDomain);
                                        }
                                        else
                                        {
                                            netCred = new NetworkCredential(EmailId, Password1);
                                        }
                                    }
                                    else
                                    {
                                        LogExceptions(strLogFilePath, null, "1");
                                        mm.From = new MailAddress(strSmtpUid, strFromName);
                                        //netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                                        netCred = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                                    }
                                }
                                else
                                {
                                   
                                    mm.From = new MailAddress(strSmtpUid, strFromName);
                                    //netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                                    netCred = new NetworkCredential(SmtpEmailValue, strSmtpPwd, Domain);
                                }
                            }
                            else
                            {
                               
                                mm.From = new MailAddress(strSmtpUid, strFromName);
                                //netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                                netCred = new NetworkCredential(strSmtpUid, strSmtpPwd, strDomain);
                            }
                        }
                        else
                        {
                            
                            mm.From = new MailAddress(strSmtpUid, strFromName);
                            //netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                            netCred = new NetworkCredential(SmtpEmailValue, strSmtpPwd, Domain);
                        }
                    }
                    else
                    {
                       
                        mm.From = new MailAddress(strSmtpUid, strFromName);
                        //netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                        netCred = new NetworkCredential(SmtpEmailValue, strSmtpPwd, Domain);
                    }
                }
                else
                {
                   
                    mm.From = new MailAddress(strSmtpUid, strFromName);
                    //netCred = new NetworkCredential(strSmtpUid, strSmtpPwd);
                    netCred = new NetworkCredential(SmtpEmailValue, strSmtpPwd, strDomain);
                }

                smtpC.Host = strSmtpServer;
                smtpC.Port = Convert.ToInt16(strSmtpPort);

                if (strEnableSSL == "Y")
                {
                    smtpC.EnableSsl = true;
                    
                }
                else
                {
                    smtpC.EnableSsl = false;
                }
                smtpC.DeliveryMethod = SmtpDeliveryMethod.Network;
               // smtpC.UseDefaultCredentials = false;
                smtpC.Credentials = netCred;
                smtpC.Send(mm);
                return true;
            }
            catch (Exception ex)
            {
                Insert_Sys_Log(ex.Message.ToString(), sName, sUser, sPwd);
                LogExceptions(strLogFilePath, ex, null);
                return false;
            }
        }

        private ItemIdType CreateDraftMessage(ExchangeServiceBinding ewsServiceBinding, MessageType emMessage)
        {
            ItemIdType iiItemid = new ItemIdType();
            CreateItemType ciCreateItemRequest = new CreateItemType();
            ciCreateItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            ciCreateItemRequest.MessageDispositionSpecified = true;
            ciCreateItemRequest.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType dfDraftsFolder = new DistinguishedFolderIdType();
            dfDraftsFolder.Id = DistinguishedFolderIdNameType.drafts;
            ciCreateItemRequest.SavedItemFolderId.Item = dfDraftsFolder;
            ciCreateItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            ciCreateItemRequest.Items.Items = new ItemType[1];
            ciCreateItemRequest.Items.Items[0] = emMessage;
            CreateItemResponseType createItemResponse = ewsServiceBinding.CreateItem(ciCreateItemRequest);
            if (createItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                //Console.WriteLine("Error Occured");
                //Console.WriteLine(createItemResponse.ResponseMessages.Items[0].MessageText);
                //eventLogFileCRM.WriteEntry(createItemResponse.ResponseMessages.Items[0].MessageText);
                LogExceptions(strLogFilePath, null, createItemResponse.ResponseMessages.Items[0].MessageText);
            }
            else
            {
                ItemInfoResponseMessageType rmResponseMessage = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                //Console.WriteLine("Item was created");
                //Console.WriteLine("Item ID : " + rmResponseMessage.Items.Items[0].ItemId.Id.ToString());
                //Console.WriteLine("ChangeKey : " + rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString());
                iiItemid.Id = rmResponseMessage.Items.Items[0].ItemId.Id.ToString();
                iiItemid.ChangeKey = rmResponseMessage.Items.Items[0].ItemId.ChangeKey.ToString();
            }

            return iiItemid;
        }

        private ItemIdType CreateAttachment(ExchangeServiceBinding ewsServiceBinding, String fnFileName, ItemIdType iiCreateItemid)
        {
            string fName = Path.GetFileName(fnFileName);
            ItemIdType iiAttachmentItemid = new ItemIdType();
            FileStream fsFileStream = new FileStream(fnFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            byte[] bdBinaryData = new byte[fsFileStream.Length];
            long brBytesRead = fsFileStream.Read(bdBinaryData, 0, (int)fsFileStream.Length);
            fsFileStream.Close();
            FileAttachmentType faFileAttach = new FileAttachmentType();
            faFileAttach.Content = bdBinaryData;
            //faFileAttach.Name = fnFileName; // Old
            faFileAttach.Name = fName;
            CreateAttachmentType amAttachmentMessage = new CreateAttachmentType();
            amAttachmentMessage.Attachments = new AttachmentType[1];
            amAttachmentMessage.Attachments[0] = faFileAttach;
            amAttachmentMessage.ParentItemId = iiCreateItemid;
            CreateAttachmentResponseType caCreateAttachmentResponse = ewsServiceBinding.CreateAttachment(amAttachmentMessage);
            if (caCreateAttachmentResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                //Console.WriteLine("Error Occured");
                //Console.WriteLine(caCreateAttachmentResponse.ResponseMessages.Items[0].MessageText);
                LogExceptions(strLogFilePath, null, caCreateAttachmentResponse.ResponseMessages.Items[0].MessageText);
            }
            else
            {
                AttachmentInfoResponseMessageType amAttachmentResponseMessage = caCreateAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType;
                //Console.WriteLine("Attachment was created");
                //Console.WriteLine("Change Key : " + amAttachmentResponseMessage.Attachments[0].AttachmentId.RootItemChangeKey.ToString());
                iiAttachmentItemid.Id = amAttachmentResponseMessage.Attachments[0].AttachmentId.RootItemId.ToString();
                iiAttachmentItemid.ChangeKey = amAttachmentResponseMessage.Attachments[0].AttachmentId.RootItemChangeKey.ToString();
            }
            return iiAttachmentItemid;
        }

        private void SendMessage(ExchangeServiceBinding ewsServiceBinding, ItemIdType iiCreateItemid)
        {
            SendItemType siSendItem = new SendItemType();
            siSendItem.ItemIds = new BaseItemIdType[1];
            siSendItem.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType siSentItemsFolder = new DistinguishedFolderIdType();
            siSentItemsFolder.Id = DistinguishedFolderIdNameType.sentitems;
            siSendItem.SavedItemFolderId.Item = siSentItemsFolder;
            siSendItem.SaveItemToFolder = true; ;
            siSendItem.ItemIds[0] = (BaseItemIdType)iiCreateItemid;
            SendItemResponseType srSendItemReponseMessage = ewsServiceBinding.SendItem(siSendItem);
            if (srSendItemReponseMessage.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Error)
            {
                //Console.WriteLine("Error Occured");
                //Console.WriteLine(srSendItemReponseMessage.ResponseMessages.Items[0].MessageText);
                LogExceptions(strLogFilePath, null, srSendItemReponseMessage.ResponseMessages.Items[0].MessageText + strSno);
            }
            else
            {
                Console.WriteLine("Message Sent");
            }
        }

        private string NewsLetterAttachment(string serName, string uName, string pWord)
        {
            try
            {
                string shellScript = strOracleRun + "  " + strOracleReportPath + strReportName + "  " +
                                            strCondParam + "=\"" + strViewName + " " + strCondValue + "\"" + "  " +
                                            "PClient=" + "\'" + strClient + "\'" + " " +
                                            "userid=" + uName + "/" + pWord + "@" + serName + "  " +
                                            "destype=file desformat=pdf batch=yes mode=BITMAP desname=" + "\"" + AppPath + strReportFileName + ".pdf" + "\"";
                //"destype=file desformat=pdf batch=yes mode=BITMAP desname=" +"\"" + @"C:\R" + "\\" + strReportFileName + ".pdf" + "\"";
                LogExceptions(strLogFilePath, null, "-------------");
                LogExceptions(strLogFilePath, null, shellScript);
                LogExceptions(strLogFilePath, null, "-------------");
                System.Diagnostics.Process process1;
                process1 = new System.Diagnostics.Process();

                string strCmdLine;
                strCmdLine = "/C " + shellScript;

                process1.StartInfo.FileName = "cmd.exe";
                process1.StartInfo.Arguments = strCmdLine;
                process1.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                process1.StartInfo.CreateNoWindow = false;
                process1.StartInfo.UseShellExecute = true;
                process1.Start();
                System.Threading.Thread.Sleep(20000);
                process1.Close();

                //string fName = @"C:\R" + "\\" + strReportFileName + ".pdf";
                fName = AppPath + strReportFileName + ".pdf";

                return fName;
            }
            catch (Exception ex)
            {
                //fName = "";
                //string serName, string uName, string pWord
                Insert_Sys_Log(ex.Message.ToString(), serName, uName, pWord);
                LogExceptions(strLogFilePath, ex, null);
                return fName;
            }
        }
        public string Encrypt(string s)
        {
            string s1 = null;
            FE_Symmetric crpto = new FE_Symmetric();
            if ((!string.IsNullOrEmpty(s)))
            {
                s1 = crpto.EncryptData("OCTaaRRMd12", s);
            }
            else
            {
                s1 = "";
            }
            return s1;
        }

        public string Decrypt(string s)
        {
            string s1 = null;
            FE_Symmetric crpto = new FE_Symmetric();
            if ((!string.IsNullOrEmpty(s)))
            {
                s1 = crpto.DecryptData("OCTaaRRMd12", s);
            }
            else
            {
                s1 = "";
            }
            return s1;
        }

        public void Insert_Sys_Log(string message, string strServerName1, string strUserName1, string strPassWord1)
        {
            try
            {
                dbcon = new OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + strServerName1 + ";User Id=" + strUserName1 + ";Password=" + strPassWord1 + ";");
                string sterr1 = null;
                string sterr2 = null;
                string sterr3 = null;
                string sterr4 = null;
                string sterr = null;
                sterr = message.Replace("'", "''");
                if (sterr.Length > 4000)
                {
                   
                    sterr1 = sterr.Substring(1, 4000);
                    if (sterr.Length > 8000)
                    {
                        
                        sterr2 = sterr.Substring(4000, 8000);
                        if (sterr.Length > 12000)
                        {
                            
                            sterr3 = sterr.Substring(8000, 12000);
                            if ((sterr.Length > 16000))
                            {
                                sterr4 = sterr.Substring(12000, 16000);
                            }
                            else
                            {
                                sterr4 = sterr.Substring(12000, sterr.Length);
                            }
                        }
                        else
                        {
                            sterr3 = sterr.Substring(8000, sterr.Length);
                            sterr4 = "";
                        }
                    }
                    else
                    {
                        sterr2 = sterr.Substring(4000, sterr.Length);
                        sterr3 = "";
                        sterr3 = "";
                        sterr4 = "";
                    }
                }
                else
                {
                    sterr1 = sterr;
                    sterr2 = "";
                    sterr3 = "";
                    sterr4 = "";
                }
                dbad.InsertCommand = new OleDbCommand();
                dbad.InsertCommand.Connection = dbcon;
                dbad.InsertCommand.CommandText = "Insert into SYS_ACTIVATE_STATUS_LOG (LINE_NO, CHANGE_REQUEST_NO,  OBJECT_TYPE, OBJECT_NAME, ERROR_TEXT, STATUS,LOG_DATE,ERROR_TEXT1, ERROR_TEXT2, ERROR_TEXT3) values ((select nvl(max(to_number(line_no)),0)+1 from SYS_ACTIVATE_STATUS_LOG),'EDGE','SERVICE','EMAIL_SERVICE','" + sterr1 + "','N',sysdate,'" + sterr2 + "','" + sterr3 + "','" + sterr4 + "')";
                if ((dbad.InsertCommand.Connection.State == ConnectionState.Closed))
                {
                    dbad.InsertCommand.Connection.Open();
                }
                dbad.InsertCommand.ExecuteNonQuery();
                if ((dbad.InsertCommand.Connection.State == ConnectionState.Open))
                {
                    dbad.InsertCommand.Connection.Close();
                }
            }
            catch (Exception ex)
            {
                LogExceptions(strLogFilePath, ex, "Insert_Sys_Log");
            }
        }
    }
}
