using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace UpdateInvoice.DBCon
{
    public class DbConnect
    {
        public void Condb()
        {
            var ds = new DataSet();

            ConnectionStringSettings pubs = ConfigurationManager.ConnectionStrings["Connstring"];  //读取配置文件 
            var conn = new SqlConnection(pubs.ConnectionString); //读取配置文件中的连接字符串
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "";  //SQL语句

            conn.Open(); //打开连接

            var sqlda = new SqlDataAdapter(cmd.CommandText, conn);
            sqlda.Fill(ds);

            conn.Close();//关闭连接
        }
    }
}
