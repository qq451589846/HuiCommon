using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace Common.Message
{
    internal static class MailConfigure
    {
        private static string _MailTemplate;
        
        /// <summary>
        /// 根據Key值獲得配置文件Value
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetSettingValue(string key) 
        {
            try{
                return ConfigurationManager.AppSettings[key];
            }
            catch(ConfigurationErrorsException ex){
                throw ex;
            }
        }
        /// <summary>
        /// Path of MailTemplate file for SendMail
        /// </summary>
        public static string MailTemplate
        {
            get
            {
                if (_MailTemplate == null || _MailTemplate == string.Empty)
                {
                    MailConfigure.LoadConfigFile();
                }
                return _MailTemplate;
            }
        }

        private static string _MailServer;
        /// <summary>
        /// The MailServer Appoint by System 
        /// </summary>
        public static string MailServer
        {
            get
            {
                if (_MailServer == null || _MailServer == string.Empty)
                {
                    MailConfigure.LoadConfigFile();
                }
                return _MailServer;
            }
        }

        private static string _MailUser;
        /// <summary>
        /// The MailUser Appoint by System 
        /// </summary>
        public static string MailUser
        {
            get
            {
                if (_MailUser == null || _MailUser == string.Empty)
                {
                    MailConfigure.LoadConfigFile();
                }
                return _MailUser;
            }
        }

        private static string _MailPassword;
        /// <summary>
        /// The MailPassword Appoint by System 
        /// </summary>
        public static string MailPassword
        {
            get
            {
                if (_MailPassword == null || _MailPassword == string.Empty)
                {
                    MailConfigure.LoadConfigFile();
                }
                return _MailPassword;
            }
        }

        private static string _MailSender;
        /// <summary>
        /// The MailSender Appoint by System 
        /// </summary>
        public static string MailSender
        {
            get
            {
                if (_MailSender == null || _MailSender == string.Empty)
                {
                    MailConfigure.LoadConfigFile();
                }
                return _MailSender;
            }
        }

        private static string _MailFontType;
        /// <summary>
        /// MailFontStype 
        /// by liuxiang 2007-10-8
        /// </summary>
        public static string MailFontType
        {
            get
            {
                return MailConfigure._MailFontType;
            }
        }

        private static void LoadConfigFile()
        {
            _MailServer = ConfigurationManager.AppSettings["MailServer"];
            _MailUser = ConfigurationManager.AppSettings["MailUser"];
            _MailPassword = ConfigurationManager.AppSettings["MailPassword"];
            _MailSender = ConfigurationManager.AppSettings["MailSender"];
            _MailTemplate = AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["MailTemplate"];
            _MailFontType = ConfigurationManager.AppSettings["MailFontType"];
        }

        static MailConfigure()
        {
            LoadConfigFile();
        }
    }
}
