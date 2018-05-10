namespace Common.Mongo
{
    public class MongoUpdateOrAddModel
    {
        /// <summary>
        /// Mongo里更新的查询条件
        /// </summary>
        public string Filter { get; set; }

        /// <summary>
        /// 需要更新的字段（如果需要新增,必须包括所有字段）
        /// </summary>
        public string Update { get; set; }
    }
}
