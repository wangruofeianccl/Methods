using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace M.MysqlHelper
{
    public static class mysqlHelper
    {
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        /// <returns></returns>
        private static string connectString()
        {
            //string M_str_sqlcon = "server=10.45.67.203;user id=root;password=123qwe123;database=vote"; //根据自己的设置
            //string M_str_sqlcon = "server=localhost;user id=root;password=123456;database=vote"; //根据自己的设置
            
            String mysqlStr = "Database=vote;Data Source=localhost;User Id=root;Password=123456;pooling=false;CharSet=utf8;port=3306";
            return mysqlStr;
        }
        ///<summary>
        /// 建立mysql数据库链接
        /// </summary>
        /// <returns></returns>
        private static MySqlConnection getMySqlCon()
        {
            string mysqlStr = connectString();
            // String mySqlCon = ConfigurationManager.ConnectionStrings["MySqlCon"].ConnectionString;
            MySqlConnection mysql = new MySqlConnection(mysqlStr);
            mysql.Open();
            return mysql;
        }
        /// <summary>
        /// 建立执行命令语句对象
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="mysql"></param>
        /// <returns></returns>
        private static MySqlCommand getSqlCommand(String sql, MySqlConnection mysql)
        {
            MySqlCommand mySqlCommand = new MySqlCommand(sql, mysql);
            //  MySqlCommand mySqlCommand = new MySqlCommand(sql);
            // mySqlCommand.Connection = mysql;
            return mySqlCommand;
        }
        /// <summary>
        /// 返回DataSet
        /// </summary>
        /// <param name="sqlStr">要执行的查询SQL语句("select * from tb_user")</param>
        /// <returns></returns>
        public static DataSet returnDataSet(string sqlStr)
        {
            try
            {
                var mycon = getMySqlCon();
                MySqlDataAdapter mda = new MySqlDataAdapter(sqlStr, mycon);
                DataSet ds = new DataSet();
                mda.Fill(ds, "table1");
                return ds;
            }
            catch
            {
                return null;
            }
           
        }
        /// <summary>
        /// 添加数据
        /// </summary>
        /// <param name="mySqlCommand"></param>
        public static int getInsert(string sql)
        {
            try
            {
                using (MySqlCommand mySqlCommand = new MySqlCommand(sql, getMySqlCon()))
                {
                    return mySqlCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }
        /// <summary>
        /// 修改数据
        /// </summary>
        /// <param name="mySqlCommand"></param>
        public static int getUpdate(string sql)
        {
            try
            {
                using (MySqlCommand mySqlCommand = new MySqlCommand(sql, getMySqlCon()))
                {
                    return mySqlCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }
        /// <summary>
        /// 删除数据
        /// </summary>
        /// <param name="mySqlCommand"></param>
        public static int getDel(string sql)
        {
            try
            {
                using (MySqlCommand mySqlCommand = new MySqlCommand(sql, getMySqlCon()))
                {
                    return mySqlCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

    }
}
