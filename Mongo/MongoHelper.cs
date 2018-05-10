using MongoDB.Driver;
using System;
using System.Collections.Generic;

namespace Common.Mongo
{
    public class MongoHelper
    {
        public MongoHelper(string mongoDBConnectionStr, string dataBaseName)
        {
            _mongoDBConnectionStr = mongoDBConnectionStr;
            _dataBaseName = dataBaseName;
        }
        public IMongoCollection<T> GetCollection<T>(string collectionName)
        {
            try
            {
                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);

                var collection = database.GetCollection<T>(collectionName);

                return collection;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool CreateOneUniqueIndex<T>(string collectionName, string indexText) where T : class
        {
            try
            {
                CreateIndexOptions createIndexOptions;

                createIndexOptions = new CreateIndexOptions();
                createIndexOptions.Unique = true;

                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);


                var collection = database.GetCollection<T>(collectionName);

                collection.Indexes.CreateOne(indexText, createIndexOptions);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool CreateOneIndex<T>(string collectionName, string indexText) where T : class
        {
            try
            {
                CreateIndexOptions createIndexOptions;

                createIndexOptions = new CreateIndexOptions();
                createIndexOptions.Unique = false;

                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);


                var collection = database.GetCollection<T>(collectionName);

                collection.Indexes.CreateOne(indexText, createIndexOptions);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public void BatchInsertMany<T>(string collectionName, List<T> dataList)
        {
            int batchCount;
            int batchSize;//将一批用户ID,分批写入MongoDB
            IMongoCollection<T> stockGameCollection;

            try
            {
                if (null == dataList || dataList.Count == 0)
                {
                    return;
                }

                batchSize = 5000;//5000一批插入MongoDB
                stockGameCollection = GetCollection<T>(collectionName);

                batchCount = dataList.Count / batchSize;
                if (dataList.Count % batchSize > 0)
                {
                    batchCount = batchCount + 1;
                }

                for (int i = 0; i < batchCount; i++)
                {

                    try
                    {
                        if (i == batchCount - 1)
                        {
                            InsertMany<T>(stockGameCollection, dataList.GetRange(i * batchSize, dataList.Count - (i * batchSize)));
                        }
                        else
                        {
                            InsertMany<T>(stockGameCollection, dataList.GetRange(i * batchSize, batchSize));
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Insert<T>(string collectionName, T data)
        {
            try
            {
                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);

                var collection = database.GetCollection<T>(collectionName);

                collection.InsertOne(data);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public void InsertMany<T>(string collectionName, List<T> dataList)
        {
            try
            {
                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);

                var collection = database.GetCollection<T>(collectionName);

                collection.InsertMany(dataList);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public void InsertMany<T>(IMongoCollection<T> collection, List<T> dataList)
        {
            try
            {
                if (null != dataList && dataList.Count > 0)
                {
                    collection.InsertMany(dataList);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public long DeleteMany<T>(string collectionName, string delFilter)
        {
            try
            {
                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);

                var collection = database.GetCollection<T>(collectionName);

                DeleteResult result = collection.DeleteMany(delFilter);

                return result.DeletedCount;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public long DeleteMany<T>(IMongoCollection<T> collection, string delFilter)
        {
            try
            {
                DeleteResult result = collection.DeleteMany(delFilter);

                return result.DeletedCount;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void UpdateMany<T>(string collectionName, string filter, string update)
        {
            try
            {
                MongoDB.Driver.UpdateOptions updateOptions;

                updateOptions = new MongoDB.Driver.UpdateOptions();
                updateOptions.IsUpsert = false;

                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);

                var collection = database.GetCollection<T>(collectionName);

                collection.UpdateMany(filter, update, updateOptions);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CreateOrUpdateOne<T>(string collectionName, string filter, string update)
        {
            try
            {
                MongoDB.Driver.UpdateOptions updateOptions;

                updateOptions = new MongoDB.Driver.UpdateOptions();

                updateOptions.IsUpsert = true;

                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);

                var collection = database.GetCollection<T>(collectionName);

                collection.UpdateOne(filter, update);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void UpdateMany<T>(IMongoCollection<T> collection, string filter, string update)
        {
            try
            {
                MongoDB.Driver.UpdateOptions updateOptions;

                updateOptions = new MongoDB.Driver.UpdateOptions();
                updateOptions.IsUpsert = false;

                collection.UpdateMany(filter, update, updateOptions);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void DropCollection(string collectionName)
        {
            try
            {
                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);
                database.DropCollection(collectionName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void BatchUpdateOrAdd<T>(string collectionName, List<MongoUpdateOrAddModel> updateOrAddDataList) where T : class
        {
            try
            {
                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);
                UpdateOneModel<T> update;

                var collection = database.GetCollection<T>(collectionName);

                List<WriteModel<T>> list = new List<WriteModel<T>>();

                foreach (var v in updateOrAddDataList)
                {
                    update = new UpdateOneModel<T>(v.Filter, v.Update);
                    update.IsUpsert = true;
                    list.Add(update);
                }

                collection.BulkWrite(list);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void BatchUpdate<T>(string collectionName, List<MongoUpdateOrAddModel> updateOrAddDataList) where T : class
        {
            try
            {
                var database = MongoUtils.CreateMongoDB(_mongoDBConnectionStr, _dataBaseName);
                UpdateOneModel<T> update;

                var collection = database.GetCollection<T>(collectionName);

                List<WriteModel<T>> list = new List<WriteModel<T>>();

                foreach (var v in updateOrAddDataList)
                {
                    update = new UpdateOneModel<T>(v.Filter, v.Update);
                    update.IsUpsert = false;
                    list.Add(update);
                }

                collection.BulkWrite(list);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private string _mongoDBConnectionStr;
        private string _dataBaseName;
    }
}
