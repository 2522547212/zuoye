using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace Excel_ClassLibrary
{
    public class Excel
    {
        //获取数据库
        public static DataTable GetDataTable(string sql, string path)
        {
            //1.构建连接数据库的字符串
            //string SConnectionString = "Provider=Microsoft.ACE.OLED.12.0;" 
            //    + "Data Source=" + path + ";"
            //    + "Extended ProPerties=Excel 8.0;HDR=Yes,IMEX=0";
            string SConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'";
             //2.实例化连接数据库
            using (OleDbConnection ole_cnn= new OleDbConnection(SConnectionString))
            {
                //打开数据库 ->Access
                ole_cnn.Open();
                //创建操作对象
                using ( OleDbCommand ole_cmd = ole_cnn.CreateCommand())
                {
                    //执行SQL语句
                    ole_cmd.CommandText = sql;
                    //返回执行的结果
                    using(OleDbDataAdapter dapter= new OleDbDataAdapter(ole_cmd))
                    {
                        //创建DataSet用以填充
                        DataSet dr = new DataSet();
                        //填充数据
                        dapter.Fill(dr);
                        //返回表格数据
                        return dr.Tables[0];
                    }
                }
            }

        }

        //更新数据
        public static int Upadate(string sql,string path)
        {
            //1.构建连接数据库的字符串
            //string SConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" 
            //    + "Data Source=" + path + ";"
            //    + "Extended ProPerties=Excel 8.0;HDR=Yes;IMEX=0";
            string SConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'";
            
             //2.实例化连接数据库
            using (OleDbConnection ole_cnn = new OleDbConnection(SConnectionString))
            {
                //打开数据库 ->Access
                ole_cnn.Open();
                //3.创建操作对象
                using (OleDbCommand ole_cmd = ole_cnn.CreateCommand())
                {
                    //4.执行SQL语句
                    ole_cmd.CommandText = sql;
                    //5.返回受影响的行，
                    return ole_cmd.ExecuteNonQuery();
                }
            }
        }

    }
}
