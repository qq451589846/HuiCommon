using System;
using System.Data;
using System.Collections;
using System.Xml;

namespace Common.Message
{

    /// <summary>
    /// MailTemplate 的摘要说明。
    /// </summary>
    public class MailTemplate
    {
        private static string MailTemplatePath = MailConfigure.MailTemplate;

        /// <summary>
        /// mailtemp
        /// </summary>
        public const string TABLE_MAIL_TEMP_NAME = "mailtemp";
        /// <summary>
        /// mailtag
        /// </summary>
        public const string TABLE_MAIL_TAG_NAME = "mailtag";
        /// <summary>
        /// mail_id
        /// </summary>
        public const string COLUMN_MAIL_ID = "mail_id";
        /// <summary>
        /// mail_name
        /// </summary>
        public const string COLUMN_MAIL_NAME = "mail_name";
        /// <summary>
        /// mail_title
        /// </summary>
        public const string COLUMN_MAIL_TITLE = "mail_title";
        /// <summary>
        /// mail_body
        /// </summary>
        public const string COLUMN_MAIL_BOBY = "mail_body";
        /// <summary>
        /// mail_sign
        /// </summary>
        public const string COLUMN_MAIL_SIGN = "mail_sign";
        /// <summary>
        /// mail_formreg
        /// </summary>
        public const string COLUMN_MAIL_FORMREG = "mail_formreg";
        /// <summary>
        /// mail_formmodel
        /// </summary>
        public const string COLUMN_MAIL_FORMMODEL = "mail_formmodel";
        /// <summary>
        /// tag_name
        /// </summary>
        public const string COLUMN_TAG_NAME = "tag_name";
        /// <summary>
        /// tag_value
        /// </summary>
        public const string COLUMN_TAG_VALUE = "tag_value";
        private string _MailTemplate;

        private XmlDocument m_xmlDoc;

        /// <summary>
        /// 构造函数
        /// </summary>
        public MailTemplate()
        {
            _MailTemplate = MailTemplate.MailTemplatePath;
            Init();
        }

        private void Init()
        {
            this.m_xmlDoc = new XmlDocument();
            this.m_xmlDoc.Load(_MailTemplate);
        }

        private DataSet InitDataSet()
        {
            DataSet ds = new DataSet();

            DataTable dt1 = new DataTable(TABLE_MAIL_TEMP_NAME);
            dt1.Columns.Add(COLUMN_MAIL_ID, typeof(int));
            dt1.Columns.Add(COLUMN_MAIL_FORMREG, typeof(int));
            dt1.Columns.Add(COLUMN_MAIL_NAME, typeof(string));
            dt1.Columns.Add(COLUMN_MAIL_TITLE, typeof(string));
            dt1.Columns.Add(COLUMN_MAIL_BOBY, typeof(string));
            dt1.Columns.Add(COLUMN_MAIL_SIGN, typeof(string));
            dt1.Columns.Add(COLUMN_MAIL_FORMMODEL, typeof(string));
            ds.Tables.Add(dt1);

            DataTable dt2 = new DataTable(TABLE_MAIL_TAG_NAME);
            dt2.Columns.Add(COLUMN_TAG_NAME, typeof(string));
            dt2.Columns.Add(COLUMN_TAG_VALUE, typeof(string));
            ds.Tables.Add(dt2);

            return ds;
        }

        /// <summary>
        /// Get all Mail templates
        /// </summary>
        /// <param name="_includeTag"></param>
        /// <returns></returns>
        public DataSet GetAllMailTemplates(bool _includeTag)
        {
            DataSet ds = InitDataSet();

            XmlNodeList objNodeList = m_xmlDoc.SelectNodes("/MailTemplates/MailTemplate");
            for (int i = 0; i < objNodeList.Count; i++)
            {
                DataRow dr = ds.Tables[TABLE_MAIL_TEMP_NAME].NewRow();
                dr[COLUMN_MAIL_ID] = int.Parse(objNodeList[i].Attributes[0].InnerText);
                dr[COLUMN_MAIL_FORMREG] = int.Parse(objNodeList[i].ChildNodes[0].InnerText);
                dr[COLUMN_MAIL_NAME] = objNodeList[i].ChildNodes[1].InnerText;
                dr[COLUMN_MAIL_TITLE] = objNodeList[i].ChildNodes[2].InnerText;
                dr[COLUMN_MAIL_BOBY] = objNodeList[i].ChildNodes[3].InnerText;
                dr[COLUMN_MAIL_SIGN] = objNodeList[i].ChildNodes[4].InnerText;
                dr[COLUMN_MAIL_FORMMODEL] = objNodeList[i].ChildNodes[5].InnerText;
                ds.Tables[TABLE_MAIL_TEMP_NAME].Rows.Add(dr);


                if (_includeTag)
                {
                    XmlNodeList temp = objNodeList[i].ChildNodes[4].ChildNodes;
                    for (int j = 0; j < temp.Count; j++)
                    {
                        DataRow dr1 = ds.Tables[TABLE_MAIL_TAG_NAME].NewRow();
                        dr1[COLUMN_TAG_NAME] = temp[j].ChildNodes[0].InnerText;
                        dr1[COLUMN_TAG_VALUE] = temp[j].ChildNodes[1].InnerText;
                        ds.Tables[TABLE_MAIL_TAG_NAME].Rows.Add(dr1);
                    }
                }
            }

            return ds;
        }

        /// <summary>
        /// Get Mail Templates from RegId
        /// </summary>
        /// <param name="_FormRegId"></param>
        /// <param name="_includeTag"></param>
        /// <returns></returns>
        public DataSet GetMailTemplatesByFormRegId(int _FormRegId, bool _includeTag)
        {
            DataSet ds = InitDataSet();

            XmlNodeList objNodeList = m_xmlDoc.SelectNodes("/MailTemplates/MailTemplate[formreg=" + _FormRegId + "]");
            for (int i = 0; i < objNodeList.Count; i++)
            {
                DataRow dr = ds.Tables[TABLE_MAIL_TEMP_NAME].NewRow();
                dr[COLUMN_MAIL_ID] = int.Parse(objNodeList[i].Attributes[0].InnerText);
                dr[COLUMN_MAIL_FORMREG] = int.Parse(objNodeList[i].ChildNodes[0].InnerText);
                dr[COLUMN_MAIL_NAME] = objNodeList[i].ChildNodes[1].InnerText;
                dr[COLUMN_MAIL_TITLE] = objNodeList[i].ChildNodes[2].InnerText;
                dr[COLUMN_MAIL_BOBY] = objNodeList[i].ChildNodes[3].InnerText;
                dr[COLUMN_MAIL_SIGN] = objNodeList[i].ChildNodes[4].InnerText;
                dr[COLUMN_MAIL_FORMMODEL] = objNodeList[i].ChildNodes[5].InnerText;
                ds.Tables[TABLE_MAIL_TEMP_NAME].Rows.Add(dr);

                if (_includeTag)
                {
                    XmlNodeList temp = objNodeList[i].ChildNodes[4].ChildNodes;
                    for (int j = 0; j < temp.Count; j++)
                    {
                        DataRow dr1 = ds.Tables[TABLE_MAIL_TAG_NAME].NewRow();
                        dr1[COLUMN_TAG_NAME] = temp[j].ChildNodes[0].InnerText;
                        dr1[COLUMN_TAG_VALUE] = temp[j].ChildNodes[1].InnerText;
                        ds.Tables[TABLE_MAIL_TAG_NAME].Rows.Add(dr1);
                    }
                }
            }

            return ds;
        }
        /// <summary>
        /// Get a mail template by a templateid
        /// </summary>
        /// <param name="_Id"></param>
        /// <param name="_includeTag"></param>
        /// <returns></returns>
        public DataSet GetMailTemplateById(int _Id, bool _includeTag)
        {
            DataSet ds = InitDataSet();

            XmlNode objNode = m_xmlDoc.SelectSingleNode("/MailTemplates/MailTemplate[@id=" + _Id + "]");
            if (objNode != null)
            {
                DataRow dr = ds.Tables[TABLE_MAIL_TEMP_NAME].NewRow();
                dr[COLUMN_MAIL_ID] = int.Parse(objNode.Attributes[0].InnerText);
                dr[COLUMN_MAIL_FORMREG] = int.Parse(objNode.ChildNodes[0].InnerText);
                dr[COLUMN_MAIL_NAME] = objNode.ChildNodes[1].InnerText;
                dr[COLUMN_MAIL_TITLE] = objNode.ChildNodes[2].InnerText;
                dr[COLUMN_MAIL_BOBY] = objNode.ChildNodes[3].InnerText;
                dr[COLUMN_MAIL_SIGN] = objNode.ChildNodes[4].InnerText;
                dr[COLUMN_MAIL_FORMMODEL] = objNode.ChildNodes[5].InnerText;
                ds.Tables[TABLE_MAIL_TEMP_NAME].Rows.Add(dr);

                if (_includeTag)
                {
                    XmlNodeList temp = objNode.ChildNodes[6].ChildNodes;
                    for (int j = 0; j < temp.Count; j++)
                    {
                        DataRow dr1 = ds.Tables[TABLE_MAIL_TAG_NAME].NewRow();
                        dr1[COLUMN_TAG_NAME] = temp[j].ChildNodes[0].InnerText;
                        dr1[COLUMN_TAG_VALUE] = temp[j].ChildNodes[1].InnerText;
                        ds.Tables[TABLE_MAIL_TAG_NAME].Rows.Add(dr1);
                    }
                }
            }

            return ds;
        }

        /// <summary>
        /// update a mail template
        /// </summary>
        /// <param name="_Id"></param>
        /// <param name="_ds"></param>
        /// <returns></returns>
        public bool UpdateMailTemplate(int _Id, DataSet _ds)
        {
            bool result = true;
            XmlNode objNode = m_xmlDoc.SelectSingleNode("/MailTemplates/MailTemplate[@id=" + _Id + "]");
            if (objNode != null && _ds.Tables[TABLE_MAIL_TEMP_NAME].Rows.Count > 0)
            {
                objNode.ChildNodes[1].InnerText = _ds.Tables[TABLE_MAIL_TEMP_NAME].Rows[0][COLUMN_MAIL_NAME].ToString();
                objNode.ChildNodes[2].InnerText = _ds.Tables[TABLE_MAIL_TEMP_NAME].Rows[0][COLUMN_MAIL_TITLE].ToString();
                objNode.ChildNodes[3].InnerText = _ds.Tables[TABLE_MAIL_TEMP_NAME].Rows[0][COLUMN_MAIL_BOBY].ToString();
                objNode.ChildNodes[4].InnerText = _ds.Tables[TABLE_MAIL_TEMP_NAME].Rows[0][COLUMN_MAIL_SIGN].ToString();
            }
            try
            {
                this.m_xmlDoc.Save(_MailTemplate);
            }
            catch
            {
                result = false;
            }

            return result;
        }
    }
}
