using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace M.AccessHelper
{
    public class AccessHelper
    {
         OleDbDataAdapter da = new OleDbDataAdapter();
         public AccessHelper()
        {

        }

        #region  Access操作

        /// 获得数据库连接
        /// </summary>
        /// <returns></returns>
        private static OleDbConnection GetDBConnection()
        {
            string basepath = Application.StartupPath + "\\Data\\MyDataBacse.accdb";
            string path = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + basepath + ";Persist Security Info=False";
            //return new OleDbConnection(ConfigurationSettings.AppSettings["Connetion"]);
            return new OleDbConnection(path);
        }
        /// <summary>
        /// 查询结果集
        /// </summary>
        /// <param name="sql">执行语句</param>
        /// <returns>返回一个DataTable对象</returns>
        public static DataTable GetDataTable(string sql)
        {
            using (OleDbConnection con = GetDBConnection())
            {
                OleDbCommand cmd = new OleDbCommand(sql, con);
                return GetDataTable(cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql"></param>
        /// <returns>return 大于0修改成功 小于等于0修改失败</returns>
        public static int UpdateDatable(string sql)
        {
            int type = -1;
            using (OleDbConnection con = GetDBConnection())
            {
                try
                {
                    OleDbCommand comm = new OleDbCommand(sql, con);
                    con.Open();
                    type = comm.ExecuteNonQuery();
                }
                catch { }
                finally
                {
                    con.Close();
                }
                return type;
            }
        }
        /// <summary>
        /// 删除记录
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static int DeleteDatable(string sql)
        {
            int type = -1;
            using (OleDbConnection con = GetDBConnection())
            {
                try
                {
                    OleDbCommand comm = new OleDbCommand(sql, con);
                    con.Open();
                    type = comm.ExecuteNonQuery();
                }
                catch { }
                finally
                {
                    con.Close();
                }
                return type;
            }
        }
        /// <summary>
        /// 新增记录
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        public static int InsertDtatTable(string[] info)
        {
            int Type = -1;
            using (OleDbConnection con = GetDBConnection())
            {
                try
                {
                    OleDbCommand comm = new OleDbCommand("insert into tb_Bill ([ID],[创建时间],[交易日期],[交易时间],[收入],[支出],[交易类型],[备注],[余额]) values ('" + info[0] + "','" + info[1] + "','" + info[2] + "','" + info[3] + "','" + info[4] + "','" + info[5] + "','" + info[6] + "','" + info[7] + "','" + info[8] + "')", con);
                    con.Open();
                    Type = comm.ExecuteNonQuery();
                }
                catch(Exception)
                {

                }
                finally
                {
                    con.Close();
                }
                return Type;
            }
        }


        /// <summary>
        /// 查询结果集
        /// </summary>
        /// <param name="cmd">执行语句的OleDbCommand命令</param>
        /// <returns>返回一个DataTable对象</returns>
        public static DataTable GetDataTable(OleDbCommand cmd)
        {
            DataSet ds = new DataSet();
            using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
            {
                try
                {
                    da.Fill(ds);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
            if (ds.Tables.Count > 0)
            {
                ds.Tables[0].DefaultView.RowStateFilter = DataViewRowState.Unchanged | DataViewRowState.Added | DataViewRowState.ModifiedCurrent | DataViewRowState.Deleted;
                return ds.Tables[0];
            }
            else
                return null;
        }

        /// <summary>
        /// 执行查询，并返回查询所返回的结果集中第一行的第一列。忽略其他列或行。
        /// </summary>
        /// <param name="sql">查询语句</param>
        /// <returns>返回结果集中第一行的第一列的object值</returns>
        public static object ExecuteScalar(string sql)
        {
            using (OleDbConnection con = GetDBConnection())
            {
                OleDbCommand cmd = new OleDbCommand(sql, con);
                return ExecuteScalar(cmd);
            }
        }

        /// <summary>
        /// 执行查询，并返回查询所返回的结果集中第一行的第一列。忽略其他列或行。
        /// </summary>
        /// <param name="cmd">查询命令</param>
        /// <returns>返回结果集中第一行的第一列的object值</returns>
        public static object ExecuteScalar(OleDbCommand cmd)
        {
            try
            {
                cmd.Connection.Open();
                object obj = cmd.ExecuteScalar();
                cmd.Connection.Close();
                return obj;
            }
            catch (Exception error)
            {
                cmd.Connection.Close();
                throw error;
            }
        }

        /// <summary>
        /// 更新数据集
        /// </summary>
        /// <param name="dt">要更新的数据集</param>
        /// <param name="insertCmd">插入SQL语句</param>
        /// <param name="updateCmd">更新SQL语句</param>
        /// <param name="deleteCmd">删除SQL语句</param>
        /// <returns></returns>
        private static int UpdateDataSet(DataTable dt, OleDbCommand insertCmd, OleDbCommand updateCmd, OleDbCommand deleteCmd)
        {
            using (OleDbDataAdapter da = new OleDbDataAdapter())
            {
                da.InsertCommand = insertCmd;
                da.UpdateCommand = updateCmd;
                da.DeleteCommand = deleteCmd;
                //da.UpdateBatchSize = 0; //UpdateBatchSize:指定可在一次批处理中执行的命令的数量,在Access不被支持。0:批大小没有限制。1:禁用批量更新。>1:更改是使用 UpdateBatchSize 操作的批处理一次性发送的。
                da.InsertCommand.UpdatedRowSource = UpdateRowSource.None;
                da.UpdateCommand.UpdatedRowSource = UpdateRowSource.None;
                da.DeleteCommand.UpdatedRowSource = UpdateRowSource.None;
                try
                {
                    int row = da.Update(dt);
                    return row;
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// 返回一个查询语句执行结果的表结构
        /// </summary>
        /// <param name="sql">查询语句,不支持复杂SQL</param>
        /// <returns></returns>
        private static DataTable GetTableSchema(string sql)
        {
            sql = sql.ToUpper();
            DataTable dt = null;
            using (OleDbConnection con = GetDBConnection())
            {
                OleDbCommand cmd = new OleDbCommand(sql, con);
                con.Open();
                using (OleDbDataReader dr = cmd.ExecuteReader(CommandBehavior.KeyInfo | CommandBehavior.SchemaOnly | CommandBehavior.CloseConnection))
                {
                    dt = dr.GetSchemaTable();
                }
            }
            return dt;
        }

        /// <summary>
        /// 根据输入的查询语句自动生成插入,更新,删除命令
        /// </summary>
        /// <param name="sql">查询语句</param>
        /// <param name="insertCmd">插入命令</param>
        /// <param name="updateCmd">更新命令</param>
        /// <param name="deleteCmd">删除命令</param>
        private static void GenerateUpdateSQL(string sql, OleDbCommand insertCmd, OleDbCommand updateCmd, OleDbCommand deleteCmd)
        {
            sql = sql.ToUpper();
            DataTable dt = GetTableSchema(sql);
            string tableName = dt.Rows[0]["BaseTableName"].ToString();
            List<OleDbParameter> updatePrimarykeys = new List<OleDbParameter>();//主键参数集合
            List<OleDbParameter> deletePrimarykeys = new List<OleDbParameter>();//主键参数集合,因为不能同时被OleDbCommand个命令引用,所以多申明一个
            List<OleDbParameter> insertFields = new List<OleDbParameter>();//字段参数集合
            List<OleDbParameter> updateFields = new List<OleDbParameter>();//字段参数集合
            string columns = string.Empty, values = "", set = "", where = "";
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["IsAutoIncrement"].ToString().Equals("False"))
                {
                    insertFields.Add(new OleDbParameter("@" + dr["BaseColumnName"].ToString(),
                                                               (OleDbType)dr["ProviderType"],
                                                               Convert.ToInt32(dr["ColumnSize"]),
                                                               dr["BaseColumnName"].ToString()));
                    updateFields.Add(new OleDbParameter("@" + dr["BaseColumnName"].ToString(),
                                                   (OleDbType)dr["ProviderType"],
                                                   Convert.ToInt32(dr["ColumnSize"]),
                                                   dr["BaseColumnName"].ToString()));

                    if (!string.IsNullOrEmpty(columns))
                        columns += ",";
                    columns += dr["BaseColumnName"].ToString();
                    if (!string.IsNullOrEmpty(values))
                        values += ",";
                    values += "@" + dr["BaseColumnName"].ToString();
                    if (!string.IsNullOrEmpty(set))
                        set += ",";
                    set += dr["BaseColumnName"].ToString() + "=@" + dr["BaseColumnName"].ToString();
                }
                if (dr["IsKey"].ToString().Equals("True"))
                {
                    updatePrimarykeys.Add(new OleDbParameter("@OLD_" + dr["BaseColumnName"].ToString(),
                                                                       (OleDbType)dr["ProviderType"],
                                                                       Convert.ToInt32(dr["ColumnSize"]),
                                                                       ParameterDirection.Input,
                                                                       Convert.ToBoolean(dr["AllowDBNull"]),
                                                                       Convert.ToByte(dr["NumericScale"]),
                                                                       Convert.ToByte(dr["NumericPrecision"]),
                                                                       dr["BaseColumnName"].ToString(), DataRowVersion.Original, null));
                    deletePrimarykeys.Add(new OleDbParameter("@OLD_" + dr["BaseColumnName"].ToString(),
                                                   (OleDbType)dr["ProviderType"],
                                                   Convert.ToInt32(dr["ColumnSize"]),
                                                   ParameterDirection.Input,
                                                   Convert.ToBoolean(dr["AllowDBNull"]),
                                                   Convert.ToByte(dr["NumericScale"]),
                                                   Convert.ToByte(dr["NumericPrecision"]),
                                                   dr["BaseColumnName"].ToString(), DataRowVersion.Original, null));
                    if (!string.IsNullOrEmpty(where))
                        where += " and ";
                    where += dr["BaseColumnName"].ToString() + "=@OLD_" + dr["BaseColumnName"].ToString();
                }
            }

            insertCmd.CommandText = string.Format("insert into {0} ({1}) values ({2})", tableName, columns, values);
            updateCmd.CommandText = string.Format("update {0} set {1} where {2}", tableName, set, where);
            deleteCmd.CommandText = string.Format("delete from {0} where {1}", tableName, where);
            insertCmd.Connection = GetDBConnection();
            updateCmd.Connection = GetDBConnection();
            deleteCmd.Connection = GetDBConnection();
            foreach (OleDbParameter pa in insertFields)
            {
                insertCmd.Parameters.Add(pa);
            }
            foreach (OleDbParameter pa in updateFields)
            {
                updateCmd.Parameters.Add(pa);
            }
            foreach (OleDbParameter pa in updatePrimarykeys)
            {
                updateCmd.Parameters.Add(pa);
            }
            foreach (OleDbParameter pa in deletePrimarykeys)
            {
                deleteCmd.Parameters.Add(pa);
            }
        }


        #endregion
    }
}
