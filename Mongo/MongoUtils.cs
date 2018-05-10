using MongoDB.Driver;
using System;
using System.Collections.Concurrent;

namespace Common.Mongo
{
    public class MongoUtils
    {
        public static IMongoDatabase CreateMongoDB(string mongoDBconnectionString, string databaseName)
        {
            string mongoDbkey;
            IMongoDatabase database;

            if (string.IsNullOrEmpty(mongoDBconnectionString))
            {
                throw new Exception("MongoDBconnectionString error");
            }

            mongoDbkey = string.Format("{0}_{1}", mongoDBconnectionString, databaseName);

            if (_mongoDBDic.TryGetValue(mongoDbkey, out database))
            {
                return database;
            }
            else
            {
                database = GetDatabase(mongoDBconnectionString, databaseName);

                _mongoDBDic.TryAdd(mongoDbkey, database);

                return database;
            }
        }


        #region Private Methods
        private static IMongoDatabase GetDatabase(string mongoDBconnectionString, string databaseName)
        {
            var client = new MongoClient(mongoDBconnectionString);

            var database = client.GetDatabase(databaseName);
            return database;
        }

        #endregion Private Methods

        #region Private field

        private static ConcurrentDictionary<string, IMongoDatabase> _mongoDBDic = new ConcurrentDictionary<string, IMongoDatabase>();

        #endregion Private field
    }
}
