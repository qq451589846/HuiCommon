using Common.Singleton;
using System;
using log4net;

namespace Common.Logging
{
    public class LogHelper : SingletonBase<LogHelper>
    {
        private ILog log;

        private LogHelper()
        {
            log = LogManager.GetLogger("DcDiagnoseStocks");
        }

        public void Debug(string classname, object message)
        {
            log.Debug(string.Format("Class:{0},Message:{1}", classname, message));
        }

        public void Debug(object message)
        {
            log.Debug(message);
        }

        public void Error(string classname, object message, Exception ex)
        {
            log.Error(string.Format("Class:{0},Message:{1}", classname, message), ex);
        }

        public void Error(object message, Exception exception)
        {
            log.Error(message, exception);
        }

        public void Error(object message)
        {
            log.Error(message);
        }

        public void Fatal(string classname, object message)
        {
            log.Fatal(string.Format("Class:{0},Message:{1}", classname, message));
        }

        public void Fatal(object message)
        {
            log.Fatal(message);
        }

        public void Info(string classname, object message)
        {
            log.Info(string.Format("Class:{0},Message:{1}", classname, message));
        }

        public void Info(object message)
        {
            log.Info(message);
        }

        public void Warn(string classname, object message)
        {
            log.Warn(string.Format("Class:{0},Message:{1}", classname, message));
        }

        public void Warn(object message)
        {
            log.Warn(message);
        }
    }
}
