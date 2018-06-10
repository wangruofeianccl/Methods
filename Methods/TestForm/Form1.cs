using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using  M.ExcelHelper;
using System.Data.OracleClient;
using Oracle.ManagedDataAccess.Client;
using M.OracleHelper;
using M.FilesHelper;

namespace TestForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string oradb = "Data Source=(DESCRIPTION="
             + "(ADDRESS=(PROTOCOL=TCP)(HOST=MyComputerName)(PORT=1521))"
             + "(CONNECT_DATA=(SERVICE_NAME=DemoDB)));"
             + "User Id=SYSTEM;Password=************;";
        string _connstring = "Data Source=(DESCRIPTION="
             + "(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.252.42)(PORT=1521))"
             + "(CONNECT_DATA=(SERVICE_NAME=ORCL)));"
             + "User Id=admin;Password=renda#weixin;"
             + "Provider=OraOLEDB.Oracle;";
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connString = "Provider=MSDAORA.1;User ID=admin;Password=renda#weixin;Data Source=(DESCRIPTION = (ADDRESS_LIST= (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.252.42)(PORT = 1521))) (CONNECT_DATA = (SERVICE_NAME = ORCL)))";
            OleDbConnection conn = new OleDbConnection(_connstring);
            try
            {
                string sqlStr="select * from TB_Users";
                string sqlstr1 = "select * from TB_USERS";
                
                conn.Open();
                OleDbCommand comm = new OleDbCommand(sqlstr1, conn);
                OleDbDataReader dr = comm.ExecuteReader();
                while (dr.Read())
                {
                    MessageBox.Show(dr.GetString(1) + "   " + dr.GetString(2));

                }
                MessageBox.Show(conn.State.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                conn.Close();
            }
            
            //try
            //{
                
            //    OracleConnection conn = new OracleConnection(_connstring);
            //    conn.Open();
            //    if (conn.State == ConnectionState.Broken)
            //    {
                   
            //        conn.Close();
            //        conn.Open();
            //        MessageBox.Show("数据库链接成功");
            //    }
            //    string sql = " SELECT * FROM ADMIN.TB_Users;"; // DemoOP是表T_TEST的user
            //    OracleCommand cmd = new OracleCommand(sql, conn);
            //    cmd.CommandType = CommandType.Text;
            //    DataSet ds = new DataSet();
            //    OracleDataAdapter da = new OracleDataAdapter();
            //    da.SelectCommand = cmd;
            //    da.Fill(ds);

            //    conn.Dispose();

            //}catch(Exception ex){
            //    MessageBox.Show(ex.Message);
            //}
            //finally
            //{
               
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sql = " SELECT * FROM TB_Users;";
            OracleHelper helper = new OracleHelper();
           var data = helper.ReturnDataReader(sql);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string connString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.252.42)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ORCL)));Persist Security Info=True;User ID=admin;Password=renda#weixin;";
                Oracle.ManagedDataAccess.Client.OracleConnection con = new Oracle.ManagedDataAccess.Client.OracleConnection(connString);

                con.Open();
                string sql = "SELECT * FROM TB_USERS"; // DemoOP是表T_TEST的user
                Oracle.ManagedDataAccess.Client.OracleCommand cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(sql, con);
                cmd.CommandType = CommandType.Text;
                DataSet ds = new DataSet();
                Oracle.ManagedDataAccess.Client.OracleDataAdapter da = new Oracle.ManagedDataAccess.Client.OracleDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(ds);
                var a = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show( ex.ToString());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
             string sql = "SELECT * FROM TB_USERS";
           // OracleHelper helper = new OracleHelper();
           //var data = helper.ReturnDataSet(sql, "myTable");
             M.OracleHelper.OracleHelper helper = new M.OracleHelper.OracleHelper();
            
             var data = helper.ReturnDataSet(sql, "m_Table");
           var a = 0;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string path = @"E:\icon\icon";
            var arr= FilesHelper.openFiles(path);
            var a = 0;
        }
    }
}
