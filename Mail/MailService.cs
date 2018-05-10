using Common.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace Common.Message
{
    /// <summary>
    /// MailService 的摘要说明。
    /// </summary>
    public class MailService
    {
        private static readonly string MailSenderConfig = MailConfigure.MailSender;
        private static readonly string MailServerConfig = MailConfigure.MailServer;
        private static readonly string MailUserConfig = MailConfigure.MailUser;
        private static readonly string MailPasswordConfig = MailConfigure.MailPassword;

        /// <summary>
        /// 根據Key值獲得配置文件Value
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetSettingValue(string key)
        {
            return MailConfigure.GetSettingValue(key);
        }

        /// <summary>
        /// 使用邮件模板发送邮件
        /// </summary>
        /// <param name="_arrMailAdds"></param>
        /// <param name="_mailTemplateId"></param>
        /// <param name="_coTags"></param>
        /// <param name="_coTagstitle"></param>
        /// <param name="IsHtml">是否是HTML格式</param>
        public static void MailSender(string[] _arrMailAdds, int _mailTemplateId, NameValueCollection _coTags, NameValueCollection _coTagstitle, bool IsHtml)
        {
            string[] to = _arrMailAdds;
            string from = MailService.MailSenderConfig;

            MailTemplate cls = new MailTemplate();
            DataSet ds = cls.GetMailTemplateById(_mailTemplateId, false);

            string subject = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_TITLE].ToString();
            string body = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_BOBY].ToString();
            string sign = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_SIGN].ToString();

            for (int i = 0; i < _coTags.Count; i++)
            {
                //替换邮件标题里的标签
                subject = subject.Replace(_coTags.AllKeys[i], _coTags[i]);

                //替换邮件模板里的标签
                body = body.Replace(_coTags.AllKeys[i], _coTags[i]);
            }
            if (IsHtml)
            {
                string bfiframe = body.Substring(0, body.IndexOf("<html"));
                string mdiframe = body.Substring(body.IndexOf("<html"), body.IndexOf("</html>") - bfiframe.Length);
                string afiframe = body.Substring(body.IndexOf("</html>"));
                string newbody = (bfiframe.Replace(" ", "&nbsp;")).Replace("\n", "<br>") + mdiframe + (afiframe.Replace(" ", "&nbsp;")).Replace("\n", "<br>");
                MailSenderBasic(to, from, subject, newbody);
            }

        }

        /// <summary>
        /// 使用邮件模板发送邮件
        /// </summary>
        /// <param name="_arrMailAdds">郵件地址</param>
        /// <param name="_mailTemplateId">模板Id</param>
        /// <param name="_coTags"></param>
        public static void MailSender(string[] _arrMailAdds, int _mailTemplateId, NameValueCollection _coTags)
        {
            string[] to = _arrMailAdds;
            string from = MailService.MailSenderConfig;

            MailTemplate cls = new MailTemplate();
            DataSet ds = cls.GetMailTemplateById(_mailTemplateId, false);

            string subject = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_TITLE].ToString();
            string body = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_BOBY].ToString();
            string sign = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_SIGN].ToString();

            for (int i = 0; i < _coTags.Count; i++)
            {
                //替换邮件标题里的标签
                subject = subject.Replace(_coTags.AllKeys[i], _coTags[i]);

                //替换邮件模板里的标签
                body = body.Replace(_coTags.AllKeys[i], _coTags[i]);
            }

            body = ((body.Replace(" ", "&nbsp;")).Replace("^", " ")).Replace("\n", "<br>");


            body += "<p><p><p>" + sign;
            MailSenderBasic(to, from, subject, body);
        }

        /// <summary>
        /// 使用邮件模板发送邮件(html)
        /// </summary>
        /// <param name="_arrMailAdds">郵件地址</param>
        /// <param name="_mailTemplateId">模板Id</param>
        /// <param name="_coTags"></param>
        public static void MailhtmlSender(string[] _arrMailAdds, int _mailTemplateId, NameValueCollection _coTags)
        {
            string[] to = _arrMailAdds;
            string from = MailService.MailSenderConfig;

            MailTemplate cls = new MailTemplate();
            DataSet ds = cls.GetMailTemplateById(_mailTemplateId, false);

            string subject = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_TITLE].ToString();
            string body = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_BOBY].ToString();
            string sign = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_SIGN].ToString();

            for (int i = 0; i < _coTags.Count; i++)
            {
                //替换邮件标题里的标签
                subject = subject.Replace(_coTags.AllKeys[i], _coTags[i]);

                //替换邮件模板里的标签
                body = body.Replace(_coTags.AllKeys[i], _coTags[i]);
            }

            body = (body.Replace("^", " ")).Replace("\n", "<br>");


            body += "<p><p><p>" + sign;
            MailSenderBasic(to, from, subject, body);
        }

        /// <summary>
        /// 使用邮件模板发送邮件(html)
        /// </summary>
        /// <param name="_arrMailAdds">郵件地址</param>
        /// <param name="_arrMailCC">抄送郵件地址</param>
        /// <param name="_mailTemplateId">模板Id</param>
        /// <param name="_coTags"></param>
        public static void MailhtmlSender(string[] _arrMailAdds, string[] _arrMailCC, int _mailTemplateId, NameValueCollection _coTags)
        {
            string[] to = _arrMailAdds;
            string[] cc = _arrMailCC;
            string from = MailService.MailSenderConfig;

            MailTemplate cls = new MailTemplate();
            DataSet ds = cls.GetMailTemplateById(_mailTemplateId, false);

            string subject = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_TITLE].ToString();
            string body = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_BOBY].ToString();
            string sign = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_SIGN].ToString();

            for (int i = 0; i < _coTags.Count; i++)
            {
                //替换邮件标题里的标签
                subject = subject.Replace(_coTags.AllKeys[i], _coTags[i]);

                //替换邮件模板里的标签
                body = body.Replace(_coTags.AllKeys[i], _coTags[i]);
            }

            body = (body.Replace("^", " ")).Replace("\n", "<br>");


            body += "<p><p><p>" + sign;
            MailSenderBasic(to, cc, from, subject, body);
        }

        /// <summary>
        /// 使用邮件模板发送邮件(html)
        /// </summary>
        /// <param name="_arrMailAdds">郵件地址</param>
        /// <param name="_mailTemplateId">模板Id</param>
        /// <param name="_coTags"></param>
        public static void MailhtmlSenderForSubject(string[] _arrMailAdds, int _mailTemplateId, NameValueCollection _coTags)
        {
            string[] to = _arrMailAdds;
            string from = MailService.MailSenderConfig;

            MailTemplate cls = new MailTemplate();
            DataSet ds = cls.GetMailTemplateById(_mailTemplateId, false);

            string subject = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_TITLE].ToString();
            string body = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_BOBY].ToString();
            string sign = ds.Tables[MailTemplate.TABLE_MAIL_TEMP_NAME].Rows[0][MailTemplate.COLUMN_MAIL_SIGN].ToString();

            //替换邮件模板里的标签
            for (int i = 0; i < _coTags.Count; i++)
            {
                subject = subject.Replace(_coTags.AllKeys[i], _coTags[i]);
                body = body.Replace(_coTags.AllKeys[i], _coTags[i]);
            }

            body = (body.Replace("^", " ")).Replace("\n", "<br>");
            subject = (subject.Replace("^", " ")).Replace("\n", "<br>");

            body += "<p><p><p>" + sign;
            MailSenderBasic(to, from, subject, body);
        }

        #region MailSenderBasic
        /// <summary>
        /// 文檔說明:邮件发送:基础方法
        /// 作    者:Johnny Jiang
        /// ?建日期:2010/07/05
        /// 修改日期:2010/07/05
        /// 修改??:
        ///         □2010/07/05
        /// </summary>
        /// <param name="_To">收件人地址(字符串数组)</param>
        /// <param name="_From">发件人地址(默认为web.config中配置的发件人)</param>
        /// <param name="_Subject">主题</param>
        /// <param name="_Body">邮件内容</param>     
        public static void MailSenderBasic(string[] _To, string _From, string _Subject, string _Body)
        {
            MailSenderBasic(_To, null, _From, _Subject, _Body);
        }

        /// <summary>
        /// 文檔說明:邮件发送:基础方法
        /// 作    者:Johnny Jiang
        /// ?建日期:2010/07/05
        /// 修改日期:2010/07/05
        /// 修改??:
        ///         □2010/07/05
        /// </summary>
        /// <param name="_To">收件人地址(字符串数组)</param>
        /// <param name="_CC">抄送人地址(字符串数组)</param>
        /// <param name="_From">发件人地址(默认为web.config中配置的发件人)</param>
        /// <param name="_Subject">主题</param>
        /// <param name="_Body">邮件内容</param>     
        public static void MailSenderBasic(string[] _To, string[] _Cc, string _From, string _Subject, string _Body)
        {
            string errMessage = string.Empty;
            try
            {
                if (_From == null || _From.Length == 0)
                {
                    _From = MailService.MailSenderConfig;
                }

                System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(MailService.MailServerConfig, 25);
                System.Net.NetworkCredential nwc = new System.Net.NetworkCredential(MailService.MailUserConfig, MailService.MailPasswordConfig);
                smtpClient.Credentials = nwc;
                smtpClient.EnableSsl = false;
                errMessage = "begin send mailMessage________";
                errMessage += "From:" + _From;
                errMessage += "________To[0]:" + _To[0];

                _Body = _Body ?? string.Empty;
                if (_Body != "")
                {
                    _Body = "<p style=\"" + MailConfigure.MailFontType + "\">" + _Body + "</p>";
                }

                System.Net.Mail.MailMessage mailMessage = null;
                if (_To != null && _To.Length>0)
                {
                    if (!String.IsNullOrEmpty(_To[0]))
                    {
                        mailMessage = new System.Net.Mail.MailMessage(MailService.MailSenderConfig, _To[0], _Subject, _Body);

                        for (int i = 1; i < _To.Length; i++)
                        {
                            if (!String.IsNullOrEmpty(_To[i]))
                            {
                                mailMessage.To.Add(_To[i]);
                                errMessage += "________To[" + i.ToString() + "]:" + _To[0];
                            }
                        }

                        if (_Cc != null && _Cc.Length > 0)
                        {
                            for (int i = 0; i < _Cc.Length; i++)
                            {
                                if (!String.IsNullOrEmpty(_Cc[i]))
                                {
                                    mailMessage.CC.Add(_Cc[i]);
                                    errMessage += "________Cc[" + i.ToString() + "]:" + _Cc[0];
                                }
                            }
                        }

                        mailMessage.BodyEncoding = Encoding.UTF8;
                        mailMessage.IsBodyHtml = true;
                        mailMessage.Priority = System.Net.Mail.MailPriority.Normal;

                        smtpClient.Send(mailMessage);
                    }
                }
  
            }
            catch (Exception ex)
            {
                errMessage += "________" + ex.ToString();
                throw ex;
            }
        }

        /// <summary>
        /// 文檔說明:發送郵件，返回發送失敗郵件列表
        /// 作    者:Johnny Jiang
        /// ?建日期:2010/07/06
        /// 修改日期:2010/07/06
        /// 修改??:
        ///         □2010/07/06
        /// </summary>
        /// <param name="_To">收件人地址(字符串数组)</param>
        /// <param name="_From">发件人地址(默认为web.config中配置的发件人)</param>
        /// <param name="_Subject">主题</param>
        /// <param name="_Body">邮件内容</param> 
        /// <param name="failedMailList">發送失敗郵件列表</param>
        public static void MailSenderBasic(string[] _To, string _From, string _CC, string _Subject, string _Body, ref IList<string> failedMailList)
        {
            if (_To != null && _To.Length > 0)
            {
                for (int i = 0; i < _To.Length; i++)
                {
                    if (!String.IsNullOrEmpty(_To[i]))
                    {
                        if (!MailSingleSender(_To[i], _From, _CC, _Subject, _Body))
                            failedMailList.Add(_To[i]);
                    }
                }
            }
        }

        #endregion

        #region MailSingleSender
        /// <summary>
        /// 文檔說明:發送郵件，返回是否發送成功
        /// 作    者:Johnny Jiang
        /// ?建日期:2010/07/06
        /// 修改日期:2010/07/06
        /// 修改??:
        ///         □2010/07/06
        /// </summary>
        /// <param name="_To">收件人地址</param>
        /// <param name="_From">发件人地址(默认为web.config中配置的发件人)</param>
        /// <param name="_CC">抄送地址</param>
        /// <param name="_Subject">主题</param>
        /// <param name="_Body">邮件内容</param> 
        /// <returns></returns>
        public static bool MailSingleSender(string _To, string _From, string _CC, string _Subject, string _Body)
        {
            bool isSend = false;
            string errMessage = string.Empty;
            try
            {
                if (_From == null || _From.Length == 0)
                {
                    _From = MailService.MailSenderConfig;
                }

                System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(MailService.MailServerConfig, 25);
                System.Net.NetworkCredential nwc = new System.Net.NetworkCredential(MailService.MailUserConfig, MailService.MailPasswordConfig);
                smtpClient.Credentials = nwc;
                smtpClient.EnableSsl = false;
                errMessage = "begin send mailMessage________";
                errMessage += "From:" + _From;
                errMessage += "________To:" + _To;

                _Body = _Body ?? string.Empty;
                if (_Body != "")
                {
                    _Body = "<p style=\"" + MailConfigure.MailFontType + "\">" + _Body + "</p>";
                }

                System.Net.Mail.MailMessage mailMessage = null;

                //mailMessage = new System.Net.Mail.MailMessage(MailService.MailSenderConfig, _To, _Subject, _Body);
                mailMessage = new MailMessage();
                mailMessage.Subject = _Subject;
                mailMessage.Body = _Body;
                mailMessage.From = new MailAddress(MailService.MailSenderConfig);
                if (!string.IsNullOrEmpty(_To))
                {
                    string[] strTO = _To.TrimEnd(';').Split(';');
                    if (strTO.Count() > 0)
                    {
                        foreach (string item in strTO)
                        {
                            try
                            {
                                mailMessage.To.Add(new MailAddress(item));
                            }
                            catch
                            {
                                LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, Receiver mail box error!", item));
                            }
                        }
                    }
                    else
                        mailMessage.To.Add(new MailAddress(_To));
                }
                if (!string.IsNullOrEmpty(_CC))
                {
                    string[] strCC = _CC.TrimEnd(';').Split(';');
                    if (strCC.Count() > 0)
                    {
                        foreach (string item in strCC)
                        {
                            try
                            {
                                mailMessage.CC.Add(new MailAddress(item));
                            }
                            catch
                            {
                                LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, CC mail box error!", item));
                            }
                        }
                    }
                    else
                        mailMessage.CC.Add(new MailAddress(_CC));
                }
                mailMessage.BodyEncoding = Encoding.UTF8;
                mailMessage.IsBodyHtml = true;
                mailMessage.Priority = System.Net.Mail.MailPriority.Normal;

                try
                {
                    smtpClient.Send(mailMessage);
                }
                catch(SmtpFailedRecipientsException e)
                {
                    for (int i = 0; i < e.InnerExceptions.Length; i++)
                    {
                        SmtpStatusCode status = e.InnerExceptions[i].StatusCode;
                        if (status == SmtpStatusCode.MailboxBusy ||
                            status == SmtpStatusCode.MailboxUnavailable)
                        {
                            LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, mail box is busy or unavailable!",
                                e.InnerExceptions[i].FailedRecipient));
                        }
                    }

                }
                catch (SmtpFailedRecipientException e)
                {
                    LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, mail box is busy or unavailable!",
                                e.FailedRecipient));
                }
                isSend = true;
            }
            catch (Exception ex)
            {
               LogHelper.Instance.Error("Type: MailSingleSender," + ex.Message);
                errMessage += "________" + ex.ToString();
            }

            return isSend;
        }

        /// <summary>
        /// 文檔說明:發送郵件，返回是否發送成功
        /// 作    者:Lily
        /// 創建日期:2014/09/03

        /// </summary>
        /// <param name="_To">收件人地址</param>
        /// <param name="_From">发件人地址(默认为web.config中配置的发件人)</param>
        /// <param name="_CC">抄送地址</param>
        /// <param name="_Subject">主题</param>
        /// <param name="_Body">邮件内容</param> 
        ///<param name="_Body">異常案件（附檔）</param> 
        ///<param name="_Body">異常員工（附檔）</param> 
        /// <returns></returns>
        public static bool MailSingleSenderForHrDataImportJob(List<string> mailList)
        {
            string _To = mailList[0];
            string _From = mailList[1];
            string _CC = mailList[2];
            string _Subject = mailList[3];
            string _Body = mailList[4];
            string _AbnormalCase = mailList[5];
            string _AbnormalEmpl = mailList[6];
            string _AbnormalCCJEmpl = mailList[7];
            string _AbnormalCCJOrg = mailList[8];
            string _AbnormalOg = mailList[9];
            string _AbnormalSupervisor = mailList[10];
            string _AbnormalUpOG = mailList[11];
            string _AbnormalIsUpORG = mailList[12];
            string _AbnormalCCJRepeat = mailList[13];
            string _AbnormalCoHead = mailList[14];
            string _AbnormalSpecialOG = mailList[15];
            string _AbnormalSpecialEmpl = mailList[16];

            bool isSend = false;
            string errMessage = string.Empty;
            string filePath = string.Empty;
            string fileName = string.Empty;
            try
            {
                if (_From == null || _From.Length == 0)
                {
                    _From = MailService.MailSenderConfig;
                }

                System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(MailService.MailServerConfig, 25);
                System.Net.NetworkCredential nwc = new System.Net.NetworkCredential(MailService.MailUserConfig, MailService.MailPasswordConfig);
                smtpClient.Credentials = nwc;
                smtpClient.EnableSsl = false;
                errMessage = "begin send mailMessage________";
                errMessage += "From:" + _From;
                errMessage += "________To:" + _To;

                _Body = _Body ?? string.Empty;
                if (_Body != "")
                {
                    _Body = "<p style=\"" + MailConfigure.MailFontType + "\">" + _Body + "</p>";
                }

                System.Net.Mail.MailMessage mailMessage = null;

                mailMessage = new MailMessage();
                mailMessage.Subject = _Subject;
                mailMessage.Body = _Body;
                mailMessage.From = new MailAddress(MailService.MailSenderConfig);
                if (!string.IsNullOrEmpty(_To))
                {
                    string[] strTO = _To.TrimEnd(';').Split(';');
                    if (strTO.Count() > 0)
                    {
                        foreach (string item in strTO)
                        {
                            try
                            {
                                mailMessage.To.Add(new MailAddress(item));
                            }
                            catch
                            {
                                LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, Receiver mail box error!", item));
                            }
                        }
                    }
                    else
                        mailMessage.To.Add(new MailAddress(_To));
                }
                if (!string.IsNullOrEmpty(_CC))
                {
                    string[] strCC = _CC.TrimEnd(';').Split(';');
                    if (strCC.Count() > 0)
                    {
                        foreach (string item in strCC)
                        {
                            try
                            {
                                mailMessage.CC.Add(new MailAddress(item));
                            }
                            catch
                            {
                                LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, CC mail box error!", item));
                            }
                        }
                    }
                    else
                        mailMessage.CC.Add(new MailAddress(_CC));
                }
                mailMessage.BodyEncoding = Encoding.UTF8;
                mailMessage.IsBodyHtml = true;
                mailMessage.Priority = System.Net.Mail.MailPriority.Normal;
                fileName = "JobResult" + DateTime.Now.ToString("yyyyMMdd") + ".xls";
                filePath = Path.GetTempPath() + fileName;

                HSSFWorkbook wk = new HSSFWorkbook();
                ICellStyle style = wk.CreateCellStyle();
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
                style.BorderTop = BorderStyle.Thin;

                #region 創建第一個Sheet
                ISheet sheet1 = wk.CreateSheet("1-異常案件");
                sheet1.SetColumnWidth(0, 20 * 256);
                IRow row1;
                ICell cell1;
                List<string> abnormalCaseList = _AbnormalCase.Split('$').ToList();
                int rowIndex1 = 0;
                foreach (var contract in abnormalCaseList)
                {
                    row1 = sheet1.CreateRow(rowIndex1);
                    cell1 = row1.CreateCell(0);
                    cell1.SetCellValue(contract);
                    cell1.CellStyle = style;
                    rowIndex1++;
                }
                #endregion

                #region 創建第二個Sheet
                ISheet sheet2 = wk.CreateSheet("2-員工部門無資料");
                sheet2.SetColumnWidth(0, 20 * 256);
                sheet2.SetColumnWidth(1, 20 * 256);
                IRow row2;
                ICell cell21;
                ICell cell22;

                List<string> abnormalEmplList = _AbnormalEmpl.Split('$').ToList();
                int rowIndex2 = 0;
                foreach (var abnormalempl in abnormalEmplList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 2)
                    {
                        row2 = sheet2.CreateRow(rowIndex2);
                        cell21 = row2.CreateCell(0);
                        cell21.SetCellValue(empl[0]);
                        cell21.CellStyle = style;
                        cell22 = row2.CreateCell(1);
                        if (rowIndex2 > 0)
                        {
                            cell22.SetCellValue(empl[1]);
                        }
                        else
                        {
                            cell22.SetCellValue(empl[1].Replace("*", " "));
                        }
                        cell22.CellStyle = style;
                    }
                    rowIndex2++;
                }
                sheet2.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
                #endregion

                #region 創建第三個Sheet
                ISheet sheet3 = wk.CreateSheet("3-員工Con-Current Jobs員工無資料");
                sheet3.SetColumnWidth(0, 20 * 256);
                IRow row3;
                ICell cell3;
                List<string> abnormalCCJEmplList = _AbnormalCCJEmpl.Split('$').ToList();
                int rowIndex3 = 0;
                foreach (var contract in abnormalCCJEmplList)
                {
                    row3 = sheet3.CreateRow(rowIndex3);
                    cell3 = row3.CreateCell(0);
                    cell3.SetCellValue(contract);
                    cell3.CellStyle = style;
                    rowIndex3++;
                }
                #endregion

                #region 創建第四個Sheet
                ISheet sheet4 = wk.CreateSheet("4-員工Con-Current Jobs部門無資料");
                sheet4.SetColumnWidth(0, 20 * 256);
                sheet4.SetColumnWidth(1, 20 * 256);
                IRow row4;
                ICell cell41;
                ICell cell42;

                List<string> abnormalCCJOrgList = _AbnormalCCJOrg.Split('$').ToList();
                int rowIndex4 = 0;
                foreach (var abnormalempl in abnormalCCJOrgList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 2)
                    {
                        row4 = sheet4.CreateRow(rowIndex4);
                        cell41 = row4.CreateCell(0);
                        cell41.SetCellValue(empl[0]);
                        cell41.CellStyle = style;
                        cell42 = row4.CreateCell(1);
                        if (rowIndex4 > 0)
                        {
                            cell42.SetCellValue(empl[1]);
                        }
                        else
                        {
                            cell42.SetCellValue(empl[1].Replace("*", " "));
                        }
                        cell42.CellStyle = style;
                    }
                    rowIndex4++;
                }
                sheet4.AddMergedRegion(new CellRangeAddress(0, 0, 0, 2));
                #endregion

                #region 創建第五個Sheet
                ISheet sheet5 = wk.CreateSheet("5-部門主管無資料");
                sheet5.SetColumnWidth(0, 20 * 256);
                sheet5.SetColumnWidth(1, 20 * 256);
                IRow row5;
                ICell cell51;
                ICell cell52;

                List<string> abnormalOgList = _AbnormalOg.Split('$').ToList();
                int rowIndex5 = 0;
                foreach (var abnormalempl in abnormalOgList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 2)
                    {
                        row5 = sheet5.CreateRow(rowIndex5);
                        cell51 = row5.CreateCell(0);
                        cell51.SetCellValue(empl[0]);
                        cell51.CellStyle = style;
                        cell52 = row5.CreateCell(1);
                        if (rowIndex5 > 0)
                        {
                            cell52.SetCellValue(empl[1]);
                        }
                        else
                        {
                            cell52.SetCellValue(empl[1].Replace("*", " "));
                        }
                        cell52.CellStyle = style;
                    }
                    rowIndex5++;
                }
                sheet5.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
                #endregion

                #region 創建第六個Sheet
                ISheet sheet6 = wk.CreateSheet("6-員工所在部門的部門主管的直屬主管無資料");
                sheet6.SetColumnWidth(0, 20 * 256);
                sheet6.SetColumnWidth(1, 20 * 256);
                sheet6.SetColumnWidth(2, 20 * 256);
                IRow row6;
                ICell cell61;
                ICell cell62;
                ICell cell63;

                List<string> abnormalSupervisorList = _AbnormalSupervisor.Split('$').ToList();
                int rowIndex6 = 0;
                foreach (var abnormalempl in abnormalSupervisorList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 3)
                    {
                        row6 = sheet6.CreateRow(rowIndex6);
                        cell61 = row6.CreateCell(0);
                        cell61.SetCellValue(empl[0]);
                        cell61.CellStyle = style;
                        cell62 = row6.CreateCell(1);
                        cell63 = row6.CreateCell(2);
                        if (rowIndex6 > 0)
                        {
                            cell62.SetCellValue(empl[1]);
                            cell63.SetCellValue(empl[2]);
                        }
                        else
                        {
                            cell62.SetCellValue(empl[1].Replace("*", " "));
                            cell63.SetCellValue(empl[2].Replace("*", " "));
                        }
                        cell62.CellStyle = style;
                        cell63.CellStyle = style;
                    }
                    rowIndex6++;
                }
                sheet6.AddMergedRegion(new CellRangeAddress(0, 0, 0, 2));
                #endregion

                #region 創建第七個Sheet
                ISheet sheet7 = wk.CreateSheet("7-部門的上層部門無資料");
                sheet7.SetColumnWidth(0, 20 * 256);
                sheet7.SetColumnWidth(1, 20 * 256);
                sheet7.SetColumnWidth(2, 20 * 256);
                IRow row7;
                ICell cell71;
                ICell cell72;

                List<string> abnormalUpOGList = _AbnormalUpOG.Split('$').ToList();
                int rowIndex7 = 0;
                foreach (var abnormalempl in abnormalUpOGList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 2)
                    {
                        row7 = sheet7.CreateRow(rowIndex7);
                        cell71 = row7.CreateCell(0);
                        cell71.SetCellValue(empl[0]);
                        cell71.CellStyle = style;
                        cell72 = row7.CreateCell(1);
                        if (rowIndex7 > 0)
                        {
                            cell72.SetCellValue(empl[1]);
                        }
                        else
                        {
                            cell72.SetCellValue(empl[1].Replace("*", " "));
                        }
                        cell72.CellStyle = style;
                    }
                    rowIndex7++;
                }
                sheet7.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
                #endregion

                #region 創建第八個Sheet
                ISheet sheet8 = wk.CreateSheet("8-部門的上層部門是本部門");
                sheet8.SetColumnWidth(0, 20 * 256);
                IRow row8;
                ICell cell8;
                List<string> abnormalIsUpORGList = _AbnormalIsUpORG.Split('$').ToList();
                int rowIndex8 = 0;
                foreach (var contract in abnormalIsUpORGList)
                {
                    row8 = sheet8.CreateRow(rowIndex8);
                    cell8 = row8.CreateCell(0);
                    cell8.SetCellValue(contract);
                    cell8.CellStyle = style;
                    rowIndex8++;
                }
                #endregion

                #region 創建第九個Sheet
                ISheet sheet9 = wk.CreateSheet("9-Con-Curretn Jobs重複的資料");
                sheet9.SetColumnWidth(0, 20 * 256);
                sheet9.SetColumnWidth(1, 20 * 256);
                IRow row9;
                ICell cell91;
                ICell cell92;

                List<string> abnormalCCJRepeatList = _AbnormalCCJRepeat.Split('$').ToList();
                int rowIndex9 = 0;
                foreach (var abnormalempl in abnormalCCJRepeatList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 2)
                    {
                        row9 = sheet9.CreateRow(rowIndex9);
                        cell91 = row9.CreateCell(0);
                        cell91.SetCellValue(empl[0]);
                        cell91.CellStyle = style;
                        cell92 = row9.CreateCell(1);
                        if (rowIndex9 > 0)
                        {
                            cell92.SetCellValue(empl[1]);
                        }
                        else
                        {
                            cell92.SetCellValue(empl[1].Replace("*", " "));
                        }
                        cell92.CellStyle = style;
                    }
                    rowIndex9++;
                }
                sheet9.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
                #endregion

                #region 創建第十個Sheet
                ISheet sheet10 = wk.CreateSheet("10-Co-head是部門正主管的資料");
                sheet10.SetColumnWidth(0, 20 * 256);
                sheet10.SetColumnWidth(1, 20 * 256);
                IRow row10;
                ICell cell101;
                ICell cell102;

                List<string> abnormalCoHeadList = _AbnormalCoHead.Split('$').ToList();
                int rowIndex10 = 0;
                foreach (var abnormalempl in abnormalCoHeadList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 2)
                    {
                        row10 = sheet10.CreateRow(rowIndex10);
                        cell101 = row10.CreateCell(0);
                        cell101.SetCellValue(empl[0]);
                        cell101.CellStyle = style;
                        cell102 = row10.CreateCell(1);
                        if (rowIndex10 > 0)
                        {
                            cell102.SetCellValue(empl[1]);
                        }
                        else
                        {
                            cell102.SetCellValue(empl[1].Replace("*", " "));
                        }
                        cell102.CellStyle = style;
                    }
                    rowIndex10++;
                }
                sheet10.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
                #endregion

                #region 創建第十一個Sheet
                ISheet sheet11 = wk.CreateSheet("11-特殊主體設定已移除部門資料");
                sheet11.SetColumnWidth(0, 20 * 256);
                sheet11.SetColumnWidth(1, 20 * 256);
                sheet11.SetColumnWidth(2, 50 * 256);
                IRow row11;
                ICell cell111;
                ICell cell112;
                ICell cell113;

                List<string> abnormalSpecialOGList = _AbnormalSpecialOG.Split('$').ToList();
                int rowIndex11 = 0;
                foreach (var abnormalempl in abnormalSpecialOGList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 3)
                    {
                        row11 = sheet11.CreateRow(rowIndex11);
                        cell111 = row11.CreateCell(0);
                        cell111.SetCellValue(empl[0]);
                        cell111.CellStyle = style;
                        cell112 = row11.CreateCell(1);
                        cell113 = row11.CreateCell(2);
                        if (rowIndex11 > 0)
                        {
                            cell112.SetCellValue(empl[1]);
                            cell113.SetCellValue(empl[2]);
                        }
                        else
                        {
                            cell112.SetCellValue(empl[1].Replace("*", " "));
                            cell113.SetCellValue(empl[2].Replace("*", " "));
                        }
                        cell112.CellStyle = style;
                        cell113.CellStyle = style;
                    }
                    rowIndex11++;
                }
                sheet11.AddMergedRegion(new CellRangeAddress(0, 0, 0, 2));
                #endregion

                #region 創建第十二個Sheet
                ISheet sheet12 = wk.CreateSheet("12-特殊主體設定已移除人員資料");
                sheet12.SetColumnWidth(0, 20 * 256);
                sheet12.SetColumnWidth(1, 20 * 256);
                sheet12.SetColumnWidth(2, 50 * 256);
                IRow row12;
                ICell cell121;
                ICell cell122;
                ICell cell123;

                List<string> abnormalSpecialEmplList = _AbnormalSpecialEmpl.Split('$').ToList();
                int rowIndex12 = 0;
                foreach (var abnormalempl in abnormalSpecialEmplList)
                {
                    var empl = abnormalempl.Split(';');
                    if (empl.Length == 3)
                    {
                        row12 = sheet12.CreateRow(rowIndex12);
                        cell121 = row12.CreateCell(0);
                        cell121.SetCellValue(empl[0]);
                        cell121.CellStyle = style;
                        cell122 = row12.CreateCell(1);
                        cell123 = row12.CreateCell(2);
                        if (rowIndex12 > 0)
                        {
                            cell122.SetCellValue(empl[1]);
                            cell123.SetCellValue(empl[2]);
                        }
                        else
                        {
                            cell122.SetCellValue(empl[1].Replace("*", " "));
                            cell123.SetCellValue(empl[2].Replace("*", " "));
                        }
                        cell122.CellStyle = style;
                        cell123.CellStyle = style;
                    }
                    rowIndex12++;
                }
                sheet12.AddMergedRegion(new CellRangeAddress(0, 0, 0, 2));
                #endregion

                //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建时不要打开该文件
                using (FileStream fs = File.OpenWrite(filePath))
                {
                    wk.Write(fs);//向打开的这个xls文件中写入并保存。  
                }
                Attachment objMailAttachment = new Attachment(filePath);
                mailMessage.Attachments.Add(objMailAttachment);//将附件附加到邮件消息对象中 
                mailMessage.Attachments.Last().ContentDisposition.FileName = fileName;
                try
                {
                    smtpClient.Send(mailMessage);
                }
                catch (SmtpFailedRecipientsException e)
                {
                    for (int i = 0; i < e.InnerExceptions.Length; i++)
                    {
                        SmtpStatusCode status = e.InnerExceptions[i].StatusCode;
                        if (status == SmtpStatusCode.MailboxBusy ||
                            status == SmtpStatusCode.MailboxUnavailable)
                        {
                            LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, mail box is busy or unavailable!",
                                e.InnerExceptions[i].FailedRecipient));
                        }
                    }

                }
                catch (SmtpFailedRecipientException e)
                {
                    LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, mail box is busy or unavailable!",
                                e.FailedRecipient));
                }
                isSend = true;
            }
            catch (Exception ex)
            {
               LogHelper.Instance.Error("Type: MailSingleSender," + ex.Message);
                errMessage += "________" + ex.ToString();
            }
            return isSend;
        }

        /// <summary>
        /// 文檔說明:發送郵件，返回是否發送成功
        /// 作    者:Johnny Jiang
        /// ?建日期:2010/07/06
        /// 修改日期:2010/07/06
        /// 修改??:
        ///         □2010/07/06
        /// </summary>
        /// <param name="_To">收件人地址</param>
        /// <param name="_From">发件人地址(默认为web.config中配置的发件人)</param>
        /// <param name="_Subject">主题</param>
        /// <param name="_Body">邮件内容</param> 
        /// <param name="_MailBG">背景圖片地址</param>
        /// <returns></returns>
        public static bool MailSingleSender(string _To, string _From, string _CC, string _Subject, string _Body, string _MailMsg, string _MailBG, ref string msg)
        {
            bool isSend = false;
            string errMessage = string.Empty;
            try
            {
                if (_From == null || _From.Length == 0)
                {
                    _From = MailService.MailSenderConfig;
                }

                System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(MailService.MailServerConfig, 25);
                System.Net.NetworkCredential nwc = new System.Net.NetworkCredential(MailService.MailUserConfig, MailService.MailPasswordConfig);
                smtpClient.Credentials = nwc;
                smtpClient.EnableSsl = false;
                errMessage = "begin send mailMessage________";
                errMessage += "From:" + _From;
                errMessage += "________To:" + _To;

                string p = "<p style=\"" + MailConfigure.MailFontType + "\"></p>";

                _Body = _Body ?? string.Empty;
                _Body = _Body.Trim();
                if (_Body != "")
                {
                    if (!_Body.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
                    {
                        _Body = string.Format("http://{0}", _Body);
                    }
                    else if (!_Body.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    {
                        _Body = string.Format("https://{0}", _Body);
                    }

                    _Body = p + _MailMsg + "  <a href='" + _Body + "'>" + _Body + "</a>";
                }
                else
                {
                    _Body = p + _MailMsg;
                }
                _Body = string.Format("<html><head></head><body background=\"{0}\" bgproperties=FIXED>{1}</body></html>", _MailBG, _Body);
                _Body = _Body + "";

                System.Net.Mail.MailMessage mailMessage = null;

                //mailMessage = new System.Net.Mail.MailMessage(MailService.MailSenderConfig, _To, _Subject, _Body);
                mailMessage = new MailMessage();
                mailMessage.Subject = _Subject;
                mailMessage.Body = _Body;
                mailMessage.From = new MailAddress(MailService.MailSenderConfig);
                if (!string.IsNullOrEmpty(_To))
                {
                    string[] strTO = _To.TrimEnd(';').Split(';');
                    if (strTO.Count() > 0)
                    {
                        foreach (string item in strTO)
                        {
                            try
                            {
                                mailMessage.To.Add(new MailAddress(item));
                            }
                            catch
                            {
                                LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, Receiver mail box error!", item));
                            }
                        }
                    }
                    else
                        mailMessage.To.Add(new MailAddress(_To));
                }
                if (!string.IsNullOrEmpty(_CC))
                {
                    string[] strCC = _CC.TrimEnd(';').Split(';');
                    if (strCC.Count() > 0)
                    {
                        foreach (string item in strCC)
                        {
                            try
                            {
                                mailMessage.CC.Add(new MailAddress(item));
                            }
                            catch
                            {
                                LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, CC mail box error!", item));
                            }
                        }
                    }
                    else
                        mailMessage.CC.Add(new MailAddress(_CC));
                }
                mailMessage.BodyEncoding = Encoding.UTF8;
                mailMessage.IsBodyHtml = true;
                mailMessage.Priority = System.Net.Mail.MailPriority.Normal;

                try
                {
                    smtpClient.Send(mailMessage);
                }
                catch (SmtpFailedRecipientsException e)
                {
                    for (int i = 0; i < e.InnerExceptions.Length; i++)
                    {
                        SmtpStatusCode status = e.InnerExceptions[i].StatusCode;
                        if (status == SmtpStatusCode.MailboxBusy ||
                            status == SmtpStatusCode.MailboxUnavailable)
                        {
                            LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, mail box is busy or unavailable!",
                                e.InnerExceptions[i].FailedRecipient));
                        }
                    }

                }
                catch (SmtpFailedRecipientException e)
                {
                    LogHelper.Instance.Error(string.Format("Failed to deliver message to {0}, mail box is busy or unavailable!",
                                e.FailedRecipient));
                }
                isSend = true;
            }
            catch (Exception ex)
            {
               LogHelper.Instance.Error("Type: MailSingleSender," + ex.Message);

                errMessage += "________" + ex.ToString();
                msg = "send email error! " + errMessage;
            }

            return isSend;
        }

        public static bool MailSingleSender(string _To, string _From, string _CC, string _Subject, string _Body, string _MailBG, ref string msg)
        {
            return MailSingleSender(_To, _From, _CC, _Subject, _Body, _MailBG, string.Empty, ref msg);
        }
        #endregion
    }
}
