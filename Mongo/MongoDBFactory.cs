using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.Mongo
{
    public static class MongoDBFactory 
    {
        private static readonly Lazy<MongoHelper> _DcManager = new Lazy<MongoHelper>(() =>
            new MongoHelper(GetAppSettings("DcManagerMongoDBConStr", string.Empty), "DcManager"));

        private static string GetAppSettings(string key, string defaultValue)
        {
            if (ConfigurationManager.AppSettings[key] == null)
            {
                return defaultValue;
            }
            else
            {
                return ConfigurationManager.AppSettings[key].ToString();
            }
        }

        public static MongoHelper DcManager
        {
            get
            {
                return _DcManager.Value;
            }
        }
    }
}
