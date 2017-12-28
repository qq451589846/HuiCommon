using Newtonsoft.Json;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Common.DataAccess
{
    public class RedisHelper
    {
        ConnectionMultiplexer redis;

        public RedisHelper(string endpoints)
        {
            var addresses = endpoints.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            ConfigurationOptions config = new ConfigurationOptions
            {
                KeepAlive = 180,
                SyncTimeout = 30 * 1000
            };

            foreach (var address in addresses)
            {
                var innerAddress = address;
                if (innerAddress.Contains('@'))
                {
                    var arr = innerAddress.Split(new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
                    config.Password = arr[0];
                    innerAddress = arr[1];
                }

                config.EndPoints.Add(innerAddress);
            }

            redis = ConnectionMultiplexer.Connect(config);
        }

        ~RedisHelper()
        {
            if (redis != null)
                redis.Dispose();
        }

        public string RPop(string key)
        {
            IDatabase db = redis.GetDatabase();
            return db.ListRightPop(key);
        }

        public void AddItemToHash<T>(string hashId, string key, T value)
        {
            IDatabase db = redis.GetDatabase();
            db.HashSet(hashId, key, JsonConvert.SerializeObject(value));
        }

        public void SetRangeInHash(string hashId, IEnumerable<KeyValuePair<string, string>> keyValuePairs)
        {
            IDatabase db = redis.GetDatabase();
            HashEntry[] entries = keyValuePairs.Select(c =>
            {
                return new HashEntry(c.Key, c.Value);
            }).ToArray();

            db.HashSet(hashId, entries);
        }

        public List<T> GetValuesFromHash<T>(string hashId)
        {
            IDatabase db = redis.GetDatabase();
            return db.HashValues(hashId).Select(c => JsonConvert.DeserializeObject<T>(c.ToString())).ToList();
        }

        public List<KeyValuePair<string, string>> ScanAllHashEntries(string hashId)
        {
            IDatabase db = redis.GetDatabase();
            var result = new List<KeyValuePair<string, string>>();

            foreach (var entry in db.HashGetAll(hashId))
            {
                result.Add(new KeyValuePair<string, string>(entry.Name, entry.Value));
            }

            return result;
        }

        public void AddItemToHash(string hashId, string key, string value)
        {
            IDatabase db = redis.GetDatabase();
            db.HashSet(hashId, key, value);
        }

        public string GetEntryFromHash(string hashId, string key)
        {
            IDatabase db = redis.GetDatabase();
            return db.HashGet(hashId, key);
        }

        public bool KeyExists(string key)
        {
            IDatabase db = redis.GetDatabase();
            return db.KeyExists(key);
        }

        public void Rename(string oldkey, string newkey)
        {
            IDatabase db = redis.GetDatabase();
            db.KeyRename(oldkey, newkey);
        }
    }
}
