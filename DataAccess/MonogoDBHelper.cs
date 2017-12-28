using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq.Expressions;
using System.Threading;
using System.Threading.Tasks;
using MongoDB.Bson;
using MongoDB.Driver;
using Common.Singleton;

namespace Common.DataAccess
{
    public class MongoDBHelper : SingletonBase<MongoDBHelper>
    {
        private readonly IMongoDatabase m_mongoDatabase;

        private MongoDBHelper()
            : this(ConfigurationManager.AppSettings["MongoServer"], ConfigurationManager.AppSettings["MongoName"])
        {
        }

        public MongoDBHelper(string conn, string databaseName)
        {
            //创建mongo服务端
            var mongoClient = this.GetClient(conn);
            //初始化mongo数据库
            this.m_mongoDatabase = this.GetMongoDatabase(mongoClient, databaseName);
        }

        /// <summary>
        /// 创建mongoServer
        /// </summary>
        /// <returns></returns>
        private IMongoClient GetClient(string conn)
        {
            return new MongoClient(new MongoUrl(conn));
        }

        private IMongoDatabase GetMongoDatabase(IMongoClient mongoClient, string databaseName)
        {
            return mongoClient.GetDatabase(databaseName);
        }

        private IMongoCollection<TDocument> GetMongoCollection<TDocument>(string collectionName)
        {
            return this.GetMongoCollection<TDocument>(this.m_mongoDatabase, collectionName);
        }

        private IMongoCollection<TDocument> GetMongoCollection<TDocument>(IMongoDatabase mongoDatabase, string collectionName)
        {
            if (string.IsNullOrEmpty(collectionName))
                throw new ArgumentNullException("collectionName");

            return mongoDatabase.GetCollection<TDocument>(collectionName);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public IAsyncCursor<TDocument> FindSync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, FindOptions<TDocument, TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.FindSync(filter, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <typeparam name="TDocument"></typeparam>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public List<TDocument> FindList<TDocument>(string collectionName, FilterDefinition<TDocument> filter, FindOptions<TDocument, TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var result = this.FindSync(collectionName, filter, options, cancellationToken);
            if (null == result)
                return new List<TDocument>(0);

            var list = new List<TDocument>();
            while (result.MoveNext())
            {
                if (null == result.Current)
                    continue;

                list.AddRange(result.Current);
            }

            return list;
        }

        public BulkWriteResult<TDocument> BulkWrite<TDocument>(string collectionName, IEnumerable<WriteModel<TDocument>> requests, BulkWriteOptions options, CancellationToken cancellationToken)
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.BulkWrite(requests, options, cancellationToken);
        }

        public async Task<BulkWriteResult<TDocument>> BulkWriteAsync<TDocument>(string collectionName, IEnumerable<WriteModel<TDocument>> requests, BulkWriteOptions options, CancellationToken cancellationToken)
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.BulkWriteAsync(requests, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public IAsyncCursor<TDocument> FindSync<TDocument>(string collectionName, Expression<Func<TDocument, bool>> filter, FindOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.FindSync(filter, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<IAsyncCursor<TDocument>> FindAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, FindOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.FindAsync(filter, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public TDocument FindOneAndDelete<TDocument>(string collectionName, FilterDefinition<TDocument> filter, FindOneAndDeleteOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.FindOneAndDelete(filter, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<TDocument> FindOneAndDeleteAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, FindOneAndDeleteOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.FindOneAndDeleteAsync(filter, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="replacement"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public TDocument FindOneAndReplace<TDocument>(string collectionName, FilterDefinition<TDocument> filter, TDocument replacement, FindOneAndReplaceOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.FindOneAndReplace(filter, replacement, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="replacement"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<TDocument> FindOneAndDeleteAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, TDocument replacement, FindOneAndReplaceOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.FindOneAndReplaceAsync(filter, replacement, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="update"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public TDocument FindOneAndUpdate<TDocument>(string collectionName, FilterDefinition<TDocument> filter, UpdateDefinition<TDocument> update, FindOneAndUpdateOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.FindOneAndUpdate(filter, update, options, cancellationToken);
        }

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="update"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<TDocument> FindOneAndUpdateAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, UpdateDefinition<TDocument> update, FindOneAndUpdateOptions<TDocument> options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.FindOneAndUpdateAsync(filter, update, options, cancellationToken);
        }

        public ReplaceOneResult ReplaceOne<TDocument>(string collectionName, FilterDefinition<TDocument> filter, TDocument repacement, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.ReplaceOne(filter, repacement, options, cancellationToken);
        }

        public ReplaceOneResult ReplaceOne<TDocument>(string collectionName, Expression<Func<TDocument, bool>> filter, TDocument repacement, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.ReplaceOne(filter, repacement, options, cancellationToken);
        }

        public async Task<ReplaceOneResult> ReplaceOneAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, TDocument repacement, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.ReplaceOneAsync(filter, repacement, options, cancellationToken);
        }

        public async Task<ReplaceOneResult> ReplaceOneAsync<TDocument>(string collectionName, Expression<Func<TDocument, bool>> filter, TDocument repacement, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.ReplaceOneAsync(filter, repacement, options, cancellationToken);
        }

        public UpdateResult UpdateMany<TDocument>(string collectionName, FilterDefinition<TDocument> filter, UpdateDefinition<TDocument> update, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.UpdateMany(filter, update, options, cancellationToken);
        }

        public async Task<UpdateResult> UpdateManyAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, UpdateDefinition<TDocument> update, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.UpdateManyAsync(filter, update, options, cancellationToken);
        }

        public UpdateResult UpdateOne<TDocument>(string collectionName, FilterDefinition<TDocument> filter, UpdateDefinition<TDocument> update, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.UpdateOne(filter, update, options, cancellationToken);
        }

        public async Task<UpdateResult> UpdateOneAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, UpdateDefinition<TDocument> update, UpdateOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.UpdateOneAsync(filter, update, options, cancellationToken);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public long Count<TDocument>(string collectionName, FilterDefinition<TDocument> filter, CountOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return collection.Count(filter, options, cancellationToken);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="collectionName"></param>
        /// <param name="filter"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<long> CountAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, CountOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var collection = this.GetMongoCollection<TDocument>(collectionName);
            return await collection.CountAsync(filter, options, cancellationToken);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TDocument"></typeparam>
        /// <param name="collectionName">mongo中表名称</param>
        /// <param name="entityList">数据</param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public void InsertBatch<TDocument>(string collectionName, IEnumerable<TDocument> entityList, InsertManyOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            col.InsertMany(entityList, options, cancellationToken);
        }

        public Task InsertManyAsync<TDocument>(string collectionName, IEnumerable<TDocument> entityList, InsertManyOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            return col.InsertManyAsync(entityList, options, cancellationToken);
        }

        public void Insert<TDocument>(string collectionName, TDocument document)
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            col.InsertOne(document);
        }

        public Task InsertAsync<TDocument>(string collectionName, TDocument document)
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            return col.InsertOneAsync(document);
        }

        /// <summary>
        /// 删除记录
        /// </summary>
        /// <param name="collectionName">Mongo表名</param>
        /// <param name="filter"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public DeleteResult DeleteMany<TDocument>(string collectionName, FilterDefinition<TDocument> filter, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            return col.DeleteMany(filter, cancellationToken);
        }

        /// <summary>
        /// 删除记录
        /// </summary>
        /// <param name="collectionName">Mongo表名</param>
        /// <param name="filter"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<DeleteResult> DeleteManyAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            return await col.DeleteManyAsync(filter, cancellationToken);
        }

        /// <summary>
        /// 删除记录
        /// </summary>
        /// <param name="collectionName">Mongo表名</param>
        /// <param name="filter"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public DeleteResult DeleteOne<TDocument>(string collectionName, FilterDefinition<TDocument> filter, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            return col.DeleteOne(filter, cancellationToken);
        }

        /// <summary>
        /// 删除记录
        /// </summary>
        /// <param name="collectionName">Mongo表名</param>
        /// <param name="filter"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<DeleteResult> DeleteOneAsync<TDocument>(string collectionName, FilterDefinition<TDocument> filter, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<TDocument>(collectionName);
            return await col.DeleteOneAsync(filter, cancellationToken);
        }

        public void RenameCollection(string oldName, string newName, RenameCollectionOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            this.m_mongoDatabase.RenameCollection(oldName, newName, options, cancellationToken);
        }

        public Task RenameCollectionAsync(string oldName, string newName, RenameCollectionOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            return this.m_mongoDatabase.RenameCollectionAsync(oldName, newName, options, cancellationToken);
        }

        public string CreateIndex(string collectionname, IndexKeysDefinition<BsonDocument> keys, CreateIndexOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<BsonDocument>(collectionname);
            return col.Indexes.CreateOne(keys, options, cancellationToken);
        }

        public async Task<string> CreateOneAsync(string collectionname, IndexKeysDefinition<BsonDocument> keys, CreateIndexOptions options = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            var col = this.GetMongoCollection<BsonDocument>(collectionname);
            return await col.Indexes.CreateOneAsync(keys, options, cancellationToken);
        }

        public void DropCollection(string collectionName, CancellationToken cancellationToken = default(CancellationToken))
        {
            this.m_mongoDatabase.DropCollection(collectionName, cancellationToken);
        }

        public async Task DropCollectionAsync(string collectionName, CancellationToken cancellationToken = default(CancellationToken))
        {
            await this.m_mongoDatabase.DropCollectionAsync(collectionName, cancellationToken);
        }

    }
}
