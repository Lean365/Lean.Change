using System;
using System.Data;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Collections.Generic;
using SqlBulkCopy;
using SqlBulkCopy.Internal.BulkCopy;
using SqlBulkCopy.Internal.Odbc;
/// <summary>
///主要包括sqlHelp数据库访问助手类 和常用的一些函数定义
///</summary>
///SqlHelp数据库访问助手
///1.public static void OpenConn()                                  //打开数据库连接
///2.public static void CloseConn()                                 //关闭数据库连接
///3.public static SqlDataReader getDataReaderValue(string sql)     //读取数据
///4.public DataSet GetDataSetValue(string sql, string tableName)   //返回DataSet
///5.public DataView GetDataViewValue(string sql)                   //返回DataView
///6.public DataTable GetDataTableValue(string sql)                 //返回DataTable
///7.public void ExecuteNonQuery(string sql)                        //执行一个SQL操作:添加、删除、更新操作
///8.public int ExecuteNonQueryCount(string sql)                    //执行一个SQL操作:添加、删除、更新操作，返回受影响的行
///9.public static object ExecuteScalar(string sql)                 //执行一条返回第一条记录第一列的SqlCommand命令
///10.public int SqlServerRecordCount(string sql)                   //返回记录数


///常用函数
///1.public static bool IsNumber(string a)                          //判断是否为数字
///2.public static string GetSafeValue(string value)                //过滤非法字符
namespace Lean.Change
{
    class Helper_Sql
    {

        ///私有属性:数据库连接字符串
        ///Data Source=(Local)          服务器地址
        ///Initial Catalog=SimpleMESDB  数据库名称
        ///User ID=sa                   数据库用户名
        ///Password=admin123456         数据库密码
       // private const string connectionString = "Data Source = 192.168.16.253; Initial Catalog = OneHerba; Persist Security Info=True;User Id==sa;Password=tac26901333;Connect Timeout = 1000";

        private static string conStr = ConfigurationManager.ConnectionStrings["UploadFs3Serv"].ConnectionString;

        /// <summary>
        /// sqlHelp 的摘要说明:数据库访问助手类
        /// Helper_Sql是从DAAB中提出的一个类，在这里进行了简化，DAAB是微软Enterprise Library的一部分，该库包含了大量大型应用程序
        /// 开发需要使用的库类。
        /// </summary>


        public void SqlHelp()
        {
            //无参构造函数
        }

        public static SqlConnection conn;

        //打开数据库连接
        public static void OpenConn()
        {
            string SqlCon = conStr;//数据库连接字符串
            conn = new SqlConnection(SqlCon);
            if (conn.State.ToString().ToLower() == "open")
            {

            }
            else
            {
                conn.Open();
            }
        }

        //关闭数据库连接
        public static void CloseConn()
        {
            if (conn.State.ToString().ToLower() == "open")
            {
                //连接打开时
                conn.Close();
                conn.Dispose();
            }
        }


        // 读取数据
        public static SqlDataReader GetDataReaderValue(string sql)
        {
            OpenConn();
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            CloseConn();
            return dr;
        }


        // 返回DataSet
        public static DataSet GetDataSetValue(string sql, string tableName)
        {
            OpenConn();
            SqlDataAdapter da;
            DataSet ds = new DataSet();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(ds, tableName);
            CloseConn();
            return ds;
        }

        //  返回DataView
        public DataView GetDataViewValue(string sql)
        {
            OpenConn();
            SqlDataAdapter da;
            DataSet ds = new DataSet();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(ds, "temp");
            CloseConn();
            return ds.Tables[0].DefaultView;
        }

        // 返回DataTable
        public static DataTable GetDataTableValue(string sql)
        {
            OpenConn();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(sql, conn);
            da.Fill(dt);
            CloseConn();
            return dt;
        }

        // 执行一个SQL操作:添加、删除、更新操作
        public void ExecuteNonQuery(string sql)
        {
            OpenConn();
            SqlCommand cmd;
            cmd = new SqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            CloseConn();
        }

        // 执行一个SQL操作:添加、删除、更新操作，返回受影响的行
        public int ExecuteNonQueryCount(string sql)
        {
            OpenConn();
            SqlCommand cmd;
            cmd = new SqlCommand(sql, conn);
            int value = cmd.ExecuteNonQuery();
            return value;
        }

        //执行一条返回第一条记录第一列的SqlCommand命令
        public object ExecuteScalar(string sql)
        {
            OpenConn();
            SqlCommand cmd;
            cmd = new SqlCommand(sql, conn);
            object value = cmd.ExecuteScalar();
            return value;
        }

        // 返回记录数
        public int SqlServerRecordCount(string sql)
        {
            OpenConn();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = sql;
            cmd.Connection = conn;
            SqlDataReader dr;
            dr = cmd.ExecuteReader();
            int RecordCount = 0;
            while (dr.Read())
            {
                RecordCount = RecordCount + 1;
            }
            CloseConn();
            return RecordCount;
        }


        ///<summary>
        ///一些常用的函数
        ///</summary>

        //判断是否为数字
        public static bool IsNumber(string a)
        {
            if (string.IsNullOrEmpty(a))
                return false;
            foreach (char c in a)
            {
                if (!char.IsDigit(c))
                    return false;
            }
            return true;
        }

        // 过滤非法字符
        public static string GetSafeValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;
            value = Regex.Replace(value, @";", string.Empty);
            value = Regex.Replace(value, @"'", string.Empty);
            value = Regex.Replace(value, @"&", string.Empty);
            value = Regex.Replace(value, @"%20", string.Empty);
            value = Regex.Replace(value, @"--", string.Empty);
            value = Regex.Replace(value, @"==", string.Empty);
            value = Regex.Replace(value, @"<", string.Empty);
            value = Regex.Replace(value, @">", string.Empty);
            value = Regex.Replace(value, @"%", string.Empty);
            return value;
        }


        /// <summary>
        /// 创建临时表存储当前需要提交的数据
        /// </summary>
        /// <param name="databaseConnectionString">数据库连接字符串</param>
        /// <param name="tableName">当前更新导入的表名</param>
        /// <returns></returns>
        public static string CreateTempTable(string databaseConnectionString, string tableName)
        {
            string tempTableName = string.Empty;
            string sql = string.Empty;
            //创建临时表存储当前需要提交的数据
            for (int i = 0; i < 20; i++)  //只是设定一个临时表的序号，如果是多台主机导入同一个数据库的时候，在数据库中可能会创建同名的临时表
            {
                sql = string.Format("SELECT * INTO {0} FROM {1} where 1=2", tableName + "_TEMP" + i, tableName);
                try
                {
                    ExecuteQuery(sql, databaseConnectionString);
                    tempTableName = tableName + "_TEMP" + i;
                    break;
                }
                catch //如果在创建当前的临时表的过程中发生了错误，如当前临时表已经存在或者被占用，则继续通过 序号i去创建下一张临时表，直到创建成功为止
                {
                    continue;
                }
            }
            return tempTableName;
        }
        /// <summary>
        /// 将临时表和当前需要更新的表的数据进行对比，根据表内的唯一标识提取出在当前表中不存在的数据，然后批量插入新增的数据
        /// </summary>
        /// <param name="databaseConnectionString">数据库连接字符串</param>
        /// /// <param name="dt">更新需要提交数据的datatable</param>
        /// <param name="tableName">当前更新导入的表名</param>
        /// <param name="tempTableName">临时表表名</param>
        /// <returns></returns>
        public static string GetUpdateData(string databaseConnectionString, DataTable dt, string tableName, string tempTableName)
        {
            string errMess = string.Empty;
            try
            {
                string sql = string.Empty;
                string found = string.Empty;
                List<string> keyList = new List<string>();
                //提取需要批量插入的数据的唯一标识列表
                sql = string.Format("select key from {0}  except select key from {1}", tempTableName, tableName);
                found = string.Format("key= '{0}'", "@key");  //生成在当前的datatable查找的语句，key为查找的关键字，当前使用@key来代替所需的关键字
                #region 批量插入数据库中不存在的数据

                keyList = new List<string>();  //用于存储当前查询的唯一标识中在临时表中存在而在数据库表中不存在的数据列表
                keyList = GetSqlAsString(sql, null, databaseConnectionString);
                DataTable appendData = dt.Clone();
                for (int i = 0; i < keyList.Count; i++)
                {
                    DataRow[] dr = dt.Select(found.Replace("@key", keyList[i]));
                    foreach (var row in dr)
                    {
                        appendData.ImportRow(row);
                    }
                }
                SqlBulkCopyByDatatable(databaseConnectionString, tableName, appendData);
                #endregion
                return errMess;
            }
            catch (Exception ex)
            {
                errMess = ex.Message;
                return errMess;
            }
        }
        /// <summary>
        /// 根据临时表更新当前导入表的数据
        /// </summary>
        /// <param name="databaseConnectionString">数据库连接字符串</param>
        /// <param name="tableName">当前更新导入的表名</param>
        /// <param name="tempTableName">临时表表名</param>
        /// <returns></returns>
        public static string UpdateTableData(string databaseConnectionString, string tableName, string tempTableName)
        {
            string errMess = string.Empty;
            try
            {
                var filedList = GetFileds(databaseConnectionString, tableName);  //获取当前操作表的字段信息
                StringBuilder sql = new StringBuilder();  //构造更新当前表的所有字段的数据信息的更新语句
                sql.Append(string.Format("update {0} set ", tableName));
                foreach (var filed in filedList)
                {
                    sql.Append(string.Format("{0} = {1}.{2},", filed, tempTableName, filed));
                }
                sql.Append("@");
                sql.Replace(",@", " ");
                sql.Append(string.Format("from {0} where tableName.key= {1}.key", tempTableName, tempTableName));  //利用唯一标识更新数据
                ExecuteQuery(sql.ToString(), databaseConnectionString);
                return errMess;
            }
            catch (Exception ex)
            {
                errMess = ex.Message;
                return errMess;
            }
        }
        /// <summary>
        /// 删除临时表
        /// </summary>
        /// <param name="databaseConnectionString">数据库连接字符串</param>
        /// <param name="tempTableName">需要删除的临时表表名</param>
        /// <returns></returns>
        public static string DropTempTable(string databaseConnectionString, string tempTableName)
        {
            string errMess = string.Empty;
            try
            {
                string sql = string.Empty;
                sql = string.Format("drop table {0}", tempTableName);
                ExecuteQuery(sql, databaseConnectionString);
                return errMess;
            }
            catch (Exception ex)
            {
                errMess = ex.Message;
                return errMess;
            }
        }

        /// <summary>
        /// 获取当前操作表的所有字段名
        /// </summary>
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="tableName">当前操作的表名</param>
        /// <returns></returns>
        private static List<string> GetFileds(string databaseConnectionString, string tableName)
        {
            List<string> _Fields = new List<string>();
            SqlConnection _Connection = new SqlConnection(databaseConnectionString);
            try
            {
                _Connection.Open();
                string key = GetTableKey(databaseConnectionString, tableName);
                string[] restrictionValues = new string[4];
                restrictionValues[0] = null; // Catalog
                restrictionValues[1] = null; // Owner
                restrictionValues[2] = tableName; // Table
                restrictionValues[3] = null; // Column

                using (DataTable dt = _Connection.GetSchema(SqlClientMetaDataCollectionNames.Columns, restrictionValues))
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        var filedName = dr["column_name"].ToString();
                        if (!filedName.Equals(key))  //不将当前表的主键添加进去
                            _Fields.Add(filedName);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                _Connection.Dispose();
            }
            return _Fields;
        }

        /// <summary>
        /// 根据表名获取主键字段
        /// </summary>
        /// <param name="databaseConnectionString">数据库连接字符串</param>
        /// <param name="TableName">表名</param>
        /// <returns>主键字段</returns>
        public static string GetTableKey(string databaseConnectionString, string TableName)
        {
            try
            {
                if (TableName == "")
                    return null;
                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT COLUMN_NAME ");
                sb.Append("FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE ");
                sb.Append("WHERE TABLE_NAME=");
                sb.Append("'" + TableName + "'");
                string key = string.Empty;
                var keyList = GetSqlAsString(sb.ToString(), null, databaseConnectionString);
                if (keyList.Count > 0)
                    key = keyList[0];
                return key;
            }
            catch
            {
                return null;
            }
        }
        public static string GetTablesKeys(string str)
        {
            return str;
        }

        /// <summary>
        /// 数据批量插入方法
        /// </summary>
        /// <param name="connectionString">目标连接字符</param>
        /// <param name="TableName">目标表</param>
        /// <param name="dt">源数据</param>
        public static void SqlBulkCopyByDatatable(string connectionString, string TableName, DataTable dt)
        {

            using (var conn = new SqlConnection(connectionString))

            using (var sbulkcopy = new System.Data.SqlClient.SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction))
            {
                try
                {
                    sbulkcopy.BulkCopyTimeout = 600;
                    sbulkcopy.DestinationTableName = TableName;
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        sbulkcopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                    }
                    sbulkcopy.WriteToServer(dt);
                }
                catch (System.Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// 以字符串列表的方式返回查询的结果集
        /// </summary>
        /// <param name="sqlText">查询语句</param>
        /// <param name="sqlParameters">事物</param>
        /// <param name="databaseConnectionString">数据库连接字符串</param>
        /// <returns></returns>
        static public List<string> GetSqlAsString(string sqlText, SqlParameter[] sqlParameters, string databaseConnectionString)
        {
            List<string> result = new List<string>();
            SqlDataReader reader;
            SqlConnection connection = new SqlConnection(databaseConnectionString);
            using (connection)
            {
                SqlCommand sqlcommand = connection.CreateCommand();
                sqlcommand.CommandText = sqlText;
                if (sqlParameters != null)
                {
                    sqlcommand.Parameters.AddRange(sqlParameters);
                }
                connection.Open();
                reader = sqlcommand.ExecuteReader();
                if (reader != null)
                {
                    while (reader.Read())
                    {
                        var re = reader.GetString(0);
                        if (!string.IsNullOrEmpty(re))
                            result.Add(re);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 执行查询结果
        /// </summary>
        /// <param name="sql">执行语句</param>
        /// <param name="databaseConnectionString">数据库连接字符串</param>
        public static void ExecuteQuery(string sql, string databaseConnectionString)
        {
            SqlConnection conn = new SqlConnection(databaseConnectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        /// <summary>
        /// 增量更新
        /// </summary>
        /// <param name="dt">更新数据表datatable形式</param>
        /// <param name="tableName">更新的数据库表名</param>
        /// <returns></returns>
        public static string UpdateExistData(DataTable dt, string tableName)
        {
            string databaseConnectionString = string.Empty;
            databaseConnectionString = "server=192.168.16.235;uid=sa;pwd=Tac26901333.;database=Sap_Data"; //初始化数据库连接字符串
            string err = string.Empty;
            string tempTableName = string.Empty;
            //创建临时表
            tempTableName = CreateTempTable(databaseConnectionString, tableName);
            //将当前数据导入到临时表中
            SqlBulkCopyByDatatable(databaseConnectionString, tempTableName, dt);
            //更新数据库，导入数据库中不存在的数据
            err += GetUpdateData(databaseConnectionString, dt, tableName, tempTableName);
            //更新数据库，更新当前数据库已经存在的数据
            err += UpdateTableData(databaseConnectionString, tableName, tempTableName);
            err += DropTempTable(databaseConnectionString, tempTableName);
            return err;
        }
        public static void UpdateDB()
        {
            Helper_Sql Helper_Sql = new Helper_Sql();
            string ecsql = "delete [oneCube].[dbo].[PP_SapEcnSub];" +
                            " insert into[oneCube].[dbo].[PP_SapEcnSub] SELECT newid(),[D_SAP_ZPABD_S001],[D_SAP_ZPABD_S002],[D_SAP_ZPABD_S003]" +
                            " ,[D_SAP_ZPABD_S004],[D_SAP_ZPABD_S005],[D_SAP_ZPABD_S006],[D_SAP_ZPABD_S007],[D_SAP_ZPABD_S008],[D_SAP_ZPABD_S009]," +
                            " [D_SAP_ZPABD_S010],[D_SAP_ZPABD_S011],[D_SAP_ZPABD_S012],[D_SAP_ZPABD_S013],[D_SAP_ZPABD_S014],[D_SAP_ZPABD_S015]" +
                            " ,[D_SAP_ZPABD_S016],[D_SAP_ZPABD_S017],''[Udf001],''[Udf002],''[Udf003],0[Udf004],0[Udf005],0[Udf006],[isDelete]," +
                            " [Remark],[Creator],[CreateTime],[Modifier],[ModifyTime] FROM[Sap_Data].[dbo].[PP_SapEcnSub];" +
                            " delete[oneCube].[dbo].[PP_SapEcn];" +
                            " insert into[oneCube].[dbo].[PP_SapEcn] SELECT[GUID],[D_SAP_ZPABD_Z001],[D_SAP_ZPABD_Z002],[D_SAP_ZPABD_Z003],[D_SAP_ZPABD_Z004]," +
                            " [D_SAP_ZPABD_Z005],[D_SAP_ZPABD_Z006],[D_SAP_ZPABD_Z007],[D_SAP_ZPABD_Z008],[D_SAP_ZPABD_Z009],[D_SAP_ZPABD_Z010],[D_SAP_ZPABD_Z011]," +
                            " [D_SAP_ZPABD_Z012],[D_SAP_ZPABD_Z013],[D_SAP_ZPABD_Z014],[D_SAP_ZPABD_Z015],[D_SAP_ZPABD_Z016],[D_SAP_ZPABD_Z017],[D_SAP_ZPABD_Z018]," +
                            " [D_SAP_ZPABD_Z019],[D_SAP_ZPABD_Z020],[D_SAP_ZPABD_Z021],[D_SAP_ZPABD_Z022],[D_SAP_ZPABD_Z023],[D_SAP_ZPABD_Z024],[D_SAP_ZPABD_Z025]," +
                            " [D_SAP_ZPABD_Z026],[D_SAP_ZPABD_Z027],''[Udf001],''[Udf002],''[Udf003],0[Udf004],0[Udf005],0[Udf006],[isDelete],[Remark],[Creator]," +
                            " [CreateTime],[Modifier],[ModifyTime] FROM[Sap_Data].[dbo].[PP_SapEcn];";
            Helper_Sql.ExecuteNonQuery(ecsql);
            string osql = " INSERT INTO [oneCube].[dbo].[PP_SapOrder] " +
                            " SELECT NEWID(),[D_SAP_COOIS_C001],[D_SAP_COOIS_C002],[D_SAP_COOIS_C003] " +
                            " ,[D_SAP_COOIS_C004],[D_SAP_COOIS_C005],[D_SAP_COOIS_C006],[D_SAP_COOIS_C007] " +
                            " ,[D_SAP_COOIS_C008],''[Udf001],''[Udf002],''[Udf003],0[Udf004],0[Udf005],0[Udf006]      " +
                            " ,[isDelete],[Remark],[Creator],[CreateTime],[Modifier],[ModifyTime] FROM [Sap_Data].[dbo].[PP_SapOrder] " +
                            " where [D_SAP_COOIS_C002] not in (select [D_SAP_COOIS_C002] from [oneCube].[dbo].[PP_SapOrder]); " +
                            " INSERT INTO [oneCube].[dbo].[PP_SapOrderSerial] " +
                            " SELECT NEWID(),[D_SAP_SER05_C001],[D_SAP_SER05_C002],[D_SAP_SER05_C003],[D_SAP_SER05_C004],''[Udf001],''[Udf002] " +
                            " ,''[Udf003],0[Udf004],0[Udf005],0[Udf006],[isDelete],[Remark],[Creator],[CreateTime],[Modifier],[ModifyTime] " +
                            " FROM [Sap_Data].[dbo].[PP_SapOrderSerial] " +
                            " where [D_SAP_SER05_C002]+[D_SAP_SER05_C003]+[D_SAP_SER05_C004] not in (select [D_SAP_SER05_C002]+[D_SAP_SER05_C003]+[D_SAP_SER05_C004]  " +
                            " from [oneCube].[dbo].[PP_SapOrderSerial]); " +
                            " INSERT INTO  [oneCube].[dbo].[PP_SapModelDest] " +
                            " SELECT  NEWID(),[D_SAP_DEST_Z001],[D_SAP_DEST_Z002],[D_SAP_DEST_Z003] " +
                            " ,''[Udf001],''[Udf002],''[Udf003],0[Udf004],0[Udf005],0[Udf006] " +
                            " ,[isDelete],[Remark],[Creator],[CreateTime],[Modifier],[ModifyTime] " +
                            " FROM [Sap_Data].[dbo].[PP_SapModelDest]where [D_SAP_DEST_Z001] not in (select [D_SAP_DEST_Z001] from [oneCube].[dbo].[PP_SapModelDest]); " +
                            " insert into oneCube.dbo.[PP_SapManhour] " +
                            " SELECT [GUID],[D_SAP_ZPBLD_Z001],[D_SAP_ZPBLD_Z002],[D_SAP_ZPBLD_Z003] " +
                            " ,[D_SAP_ZPBLD_Z004],[D_SAP_ZPBLD_Z005],[D_SAP_ZPBLD_Z006],[D_SAP_ZPBLD_Z007] " +
                            " ,[D_SAP_ZPBLD_Z008],''[Udf001],''[Udf002],''[Udf003],0[Udf004],0[Udf005] " +
                            " ,0[Udf006],[isDelete],[Remark],[Creator],[CreateTime],[Modifier],[ModifyTime] " +
                            " FROM [Sap_Data].[dbo].[PP_SapManhour] " +
                            " where [D_SAP_ZPBLD_Z001] not in (select [D_SAP_ZPBLD_Z001] from [oneCube].[dbo].[PP_SapManhour]); " +
                            " delete [oneCube].[dbo].[PP_SapMaterial]; " +
                            " insert into [oneCube].[dbo].[PP_SapMaterial] " +
                            " SELECT  [GUID],[D_SAP_ZCA1D_Z001],[D_SAP_ZCA1D_Z002],[D_SAP_ZCA1D_Z003] " +
                            " ,[D_SAP_ZCA1D_Z004],[D_SAP_ZCA1D_Z005],[D_SAP_ZCA1D_Z006],[D_SAP_ZCA1D_Z007],[D_SAP_ZCA1D_Z008] " +
                            " ,[D_SAP_ZCA1D_Z009],[D_SAP_ZCA1D_Z010],[D_SAP_ZCA1D_Z011],[D_SAP_ZCA1D_Z012] " +
                            " ,[D_SAP_ZCA1D_Z013],[D_SAP_ZCA1D_Z014],[D_SAP_ZCA1D_Z015] " +
                            " ,[D_SAP_ZCA1D_Z016],[D_SAP_ZCA1D_Z017],[D_SAP_ZCA1D_Z018] " +
                            " ,[D_SAP_ZCA1D_Z019],[D_SAP_ZCA1D_Z020],[D_SAP_ZCA1D_Z021] " +
                            " ,[D_SAP_ZCA1D_Z022],[D_SAP_ZCA1D_Z023],[D_SAP_ZCA1D_Z024] " +
                            " ,[D_SAP_ZCA1D_Z025],[D_SAP_ZCA1D_Z026],[D_SAP_ZCA1D_Z027] " +
                            " ,[D_SAP_ZCA1D_Z028],[D_SAP_ZCA1D_Z029],[D_SAP_ZCA1D_Z030] " +
                            " ,[D_SAP_ZCA1D_Z031],[D_SAP_ZCA1D_Z032],[D_SAP_ZCA1D_Z033] " +
                            " ,[D_SAP_ZCA1D_Z034],''[Udf001],''[Udf002],''[Udf003],0[Udf004],0[Udf005] " +
                            " ,0[Udf006],[isDelete],[Remark],[Creator],[CreateTime],[Modifier] " +
                            " ,[ModifyTime]  FROM [Sap_Data].[dbo].[PP_SapMaterial]; ";

            Helper_Sql.ExecuteNonQuery(osql);

        }


    }
}
