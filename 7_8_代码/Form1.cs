using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel_ClassLibrary;//引用动态链接库

namespace _7_8_代码
{
    public partial class Form1 : Form
    {
        //布尔值->判断是插入还是修改
        bool insert_into;

        public Form1()
        {
            InitializeComponent();
        }
        //查询表格
        private void button4_Click(object sender, EventArgs e)
        {
            //学号输入没有内容
            if (this.textBox5.Text == "")
            {
                //Excle地址
                var filePath = "./Excel表格.xls";
                //SQL语句
                string sql = "select 学号,姓名,班级,电话号码 from [学生信息$] where 状态='正常' ";
                //SQL语句执行
                this.dataGridView1.DataSource = Excel.GetDataTable(sql, filePath);
            }
            else
            {
                //Excle地址
                var filePath = "Excel表格.xls";
                //SQL语句
                string sql = "select 学号,姓名,班级,电话号码 from [学生信息$] where 状态='正常' " 
                 + " and "+ "学号="+ this.textBox5.Text;
                //SQL语句执行
                this.dataGridView1.DataSource = Excel.GetDataTable(sql, filePath);
            }
        }
        //创建表格按钮
        private void button3_Click(object sender, EventArgs e)
        {
            //Excle地址
            var filePath = "Excel表格.xls";
            //SQL语句
            string sql = "CREATE TABLE 学生信息([学号] INT,[姓名] VarChar,[班级] VarChar,[电话号码] VarChar,[状态] VarChar)";
            //调用更新数据库
            Excel.Upadate(sql, filePath);

        }
        //插入按钮
        private void button5_Click(object sender, EventArgs e)
        {
            this.groupBox1.Enabled = true;
            insert_into = true;
        }
        //修改按钮
        private void button7_Click(object sender, EventArgs e)
        {
            this.groupBox1.Enabled = true;
            insert_into = false;
        }        
        //提交
        private void button1_Click(object sender, EventArgs e)
        {
            //要执行的查询语句
            string sql;
            //Excle地址
            var filePath = "./Excel表格.xls";
            //判断是否是更新数据
            if (insert_into == true)
            {
                //SQL语句
                sql = "insert into [学生信息$](学号,姓名,班级,电话号码,状态) values({0},'{1}','{2}','{3}','{4}')";
                sql = string.Format(sql, this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox5.Text, "正常");
            }
            else
            {
                //SQL语句
                sql = "update [学生信息$] set 姓名='{0}',班级='{1}',电话号码='{2}',状态='正常' where 学号={3}";
                sql = string.Format(sql, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text, this.textBox1.Text);



            }

            //SQL语句执行
            this.dataGridView1.DataSource = Excel.Upadate(sql, filePath);
            //执行查询
            button4_Click(null, null);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.groupBox1.Enabled = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //要执行的查询语句
            string sql;
            //Excle地址
            var filePath = "./Excel表格.xls";
            //判断学号是否有输入
            if (this.textBox5.Text == "")
            {
                MessageBox.Show("请输入要删除的学号");
            }
            else
            {
                //构建删除的SQL语句
                sql = "UPDATE [学生信息$] set 状态='删除' where 学号={0}";
                sql = string.Format(sql, this.textBox1.Text);
                //执行SQL语句
                Excel.Upadate(sql, filePath);
                //清空查询编号
                this.textBox5.Text = "";
                //执行查询操作
                button4_Click(null, null);
            }
        }



    }
}
