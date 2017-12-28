using Common.Logging;
using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Common.DataAccess
{
    public class OracledbHelper
    {
        private static readonly Lazy<OracledbHelper> instance = new Lazy<OracledbHelper>(() => new OracledbHelper());

        public static OracledbHelper Current
        {
            get
            {
                return instance.Value;
            }
        }

        private readonly string m_connectionString;

        private readonly int m_commandTimeOut;

        public OracledbHelper(string connectionString)
        {
            this.m_connectionString = connectionString;
            this.m_commandTimeOut = Int32.Parse(ConfigurationManager.AppSettings["OracleCommandTimeOut"]);
            if (this.m_commandTimeOut <= 0)
                this.m_commandTimeOut = 30 * 1000;
        }

        public OracledbHelper()
            : this(ConfigurationManager.AppSettings["OracleDb"])
        {

        }

        private OracleConnection GetConnection(string connectionString)
        {
            return new OracleConnection(connectionString);
        }

        private OracleDataAdapter GetDataAdapter(OracleCommand cmd)
        {
            return new OracleDataAdapter(cmd);
        }

        public DataTable QueryForDataTable(string sql)
        {
            return QueryForDataTable(sql, new List<OracleParameter>());
        }

        public DataTable QueryForDataTable(string sql, List<OracleParameter> parameters, int retrycount = 3)
        {
            var retries = retrycount;//重试3次
            Exception ex = null;
            while (true)
            {
                try
                {
                    using (var conn = this.GetConnection(this.m_connectionString))
                    {
                        conn.Open();
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = sql;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandTimeout = this.m_commandTimeOut;

                            if (parameters != null && parameters.Any())
                            {
                                parameters.ForEach(c => cmd.Parameters.Add(c));
                            }

                            using (var da = this.GetDataAdapter(cmd))
                            {
                                var table = new DataTable();
                                da.Fill(table);
                                return table;
                            }
                        }
                    }
                }
                catch (Exception innerEx)
                {
                    if (retries >= 0)
                    {
                        retries--;
                        LogHelper.Instance.Warn(string.Format("Class:SecurityManager|Method:Initialize执行异常,异常原因{0}", innerEx.Message));
                    }
                    else
                    {
                        ex = innerEx;
                        break;
                    }
                }
            }

            LogHelper.Instance.Error(string.Format("Class:SecurityManager|Method:Initialize执行异常,异常原因{0}", ex.Message));
            throw ex;
        }

        public DataTable QueryForDataTable(string sql, Dictionary<string, object> parameterDict)
        {
            List<OracleParameter> parameters = new List<OracleParameter>();

            foreach (var item in parameterDict)
            {
                parameters.Add(new OracleParameter(item.Key, item.Value));
            }

            return QueryForDataTable(sql, parameters);
        }

        public object ExecuteScalar(string sql)
        {
            using (var conn = this.GetConnection(this.m_connectionString))
            {
                conn.Open();
                try
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = sql;
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandTimeout = this.m_commandTimeOut;
                        return cmd.ExecuteScalar();
                    }
                }
                finally
                {
                    if (null != conn && conn.State != ConnectionState.Closed)
                        conn.Close();
                }
            }
        }

        public void Execute<T>(string sql, PropertyInfo[] pInfos, List<T> list)
        {
            using (var conn = new OracleConnection(this.m_connectionString))
            {
                try
                {
                    conn.Open();
                    this.BatchExecute(conn, sql, pInfos, list);
                }
                finally
                {
                    if (!ConnectionState.Closed.Equals(conn.State))
                        conn.Close();
                }
            }
        }

        private void BatchExecute<T>(OracleConnection conn, string sql, PropertyInfo[] pInfos, List<T> list)
        {
            var parameterList = new List<OracleParameter>(list.Count);
            foreach (PropertyInfo pInfo in pInfos)
            {
                object[] objects = new object[list.Count];
                for (int i = 0; i < list.Count; i++)
                {
                    objects[i] = pInfo.GetValue(list[i], null);
                }

                OracleDbType dbType;
                if (typeof(decimal) == pInfo.PropertyType || typeof(decimal?) == pInfo.PropertyType)
                    dbType = OracleDbType.Decimal;
                else if (typeof(float) == pInfo.PropertyType || typeof(float?) == pInfo.PropertyType)
                    dbType = OracleDbType.Single;
                else if (typeof(double) == pInfo.PropertyType || typeof(double?) == pInfo.PropertyType)
                    dbType = OracleDbType.Double;
                else if (typeof(short) == pInfo.PropertyType || typeof(short?) == pInfo.PropertyType)
                    dbType = OracleDbType.Int16;
                else if (typeof(int) == pInfo.PropertyType || typeof(int?) == pInfo.PropertyType)
                    dbType = OracleDbType.Int32;
                else if (typeof(long) == pInfo.PropertyType || typeof(long?) == pInfo.PropertyType)
                    dbType = OracleDbType.Int64;
                else if (typeof(DateTime) == pInfo.PropertyType || typeof(DateTime?) == pInfo.PropertyType)
                    dbType = OracleDbType.Date;
                else if (typeof(char) == pInfo.PropertyType || typeof(char?) == pInfo.PropertyType)
                    dbType = OracleDbType.Char;
                else if (typeof(byte) == pInfo.PropertyType || typeof(byte?) == pInfo.PropertyType)
                    dbType = OracleDbType.Byte;
                else if (typeof(byte[]) == pInfo.PropertyType || typeof(byte?[]) == pInfo.PropertyType)
                    dbType = OracleDbType.Blob;
                else
                    dbType = OracleDbType.NVarchar2;

                parameterList.Add(new OracleParameter(pInfo.Name, dbType, objects, ParameterDirection.Input));
            }


            using (var command = conn.CreateCommand())
            {
                command.CommandText = sql;
                command.CommandType = CommandType.Text;
                command.Connection = conn;
                command.ArrayBindCount = list.Count;
                command.BindByName = true;
                command.CommandTimeout = this.m_commandTimeOut;
                parameterList.ForEach(parameter => command.Parameters.Add(parameter));
                command.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Execute non query by stored procedure and parameter list
        /// </summary>
        /// <param name="cmdText">stored procedure</param>
        /// <param name="parameters">parameter list</param>
        /// <param name="tableName"></param>
        /// <returns>execute count</returns>
        public int ExecuteNonQuery(string cmdText, List<OracleParameter> parameters, string tableName = "")
        {
            using (var conn = this.GetConnection(this.m_connectionString))
            {
                try
                {
                    conn.Open();
                    using (var command = conn.CreateCommand())
                    {
                        this.PrepareCommand(command, conn, cmdText, parameters);
                        return command.ExecuteNonQuery();
                    }
                }
                finally
                {
                    if (!ConnectionState.Closed.Equals(conn.State))
                        conn.Close();
                }
            }
        }


        public int BatchExecuteNoError(List<KeyValuePair<string, List<OracleParameter>>> cmdList)
        {
            if (cmdList == null || cmdList.Count == 0)
                return 0;

            int result = 0;
            int executeCount;
            using (var conn = this.GetConnection(this.m_connectionString))
            {
                try
                {
                    conn.Open();
                    cmdList.ForEach(cmdText =>
                    {
                        if (!string.IsNullOrEmpty(cmdText.Key))
                        {
                            using (var command = conn.CreateCommand())
                            {
                                try
                                {
                                    this.PrepareCommand(command, conn, cmdText.Key, cmdText.Value);
                                    executeCount = command.ExecuteNonQuery();
                                    if (executeCount > 0)
                                        result += executeCount;
                                }
                                catch
                                {

                                }
                            }
                        }
                    });
                }
                catch
                {

                }
                finally
                {
                    if (!ConnectionState.Closed.Equals(conn.State))
                        conn.Close();
                }

                return result;
            }
        }


        /// <summary>
        /// Prepare the execute command
        /// </summary>
        /// <param name="cmd">my sql command</param>
        /// <param name="conn">my sql connection</param>
        /// <param name="cmdText">stored procedure</param>
        /// <param name="parameterList">parameter list</param>
        protected void PrepareCommand(OracleCommand cmd, OracleConnection conn, string cmdText, List<OracleParameter> parameterList)
        {
            if (!ConnectionState.Open.Equals(conn.State))
                conn.Open();

            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            cmd.CommandType = CommandType.Text;
            cmd.CommandTimeout = this.m_commandTimeOut;
            cmd.Parameters.Clear();
            if (null == parameterList || 0 == parameterList.Count)
                return;

            parameterList.ForEach(para =>
            {
                if (!cmd.Parameters.Contains(para))
                    cmd.Parameters.Add(((ICloneable)para).Clone());
            });
        }
    }
}
