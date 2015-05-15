using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace yaojiang
{
	public partial class LuckyForm : Form
	{
		public LuckyForm()
		{
			InitializeComponent();
		}

		private void Form2_FormClosing(object sender, FormClosingEventArgs e)
		{
			Application.Exit();
		}

        private void Form2_Load(object sender, EventArgs e)
        {
           label1.Text = "兰博基尼";
           label2.Text = "特等奖";
           label7.Text = 1.ToString();   
        }
        
       
        private void shijian1(object sender, EventArgs e)
        {
            this.timer1.Start();
                 
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Rad();
        }

        private void Rad()
        {
            SqlConnection conn = null;
            try
            {
                
                string connStr = "server=.;database = Praise;uid = sa;pwd = 123456;";
                conn = new SqlConnection(connStr);
                conn.Open();
               
                //产生0-50之间的随机数
                Random rad = new Random();
                int i = rad.Next(0, 50);
                //创建sql语句
                string s = "select EmployeeID from Employee where EmployeeID =" + i;
                string q = "select Name from Employee where EmployeeID =" + i;
                string l = "select Department from Employee where EmployeeID =" + i;
                //执行命令
                SqlCommand e = new SqlCommand(s, conn);
                SqlCommand n = new SqlCommand(q, conn);
                SqlCommand d = new SqlCommand(l, conn);

                int r1 = (int)e.ExecuteScalar();
                string r2 = (string)n.ExecuteScalar();
                string r3 = (string)d.ExecuteScalar();

                textBox1.Text = r1.ToString();
                textBox2.Text = r2;
                textBox3.Text = r3;
              ;
            }
            catch (Exception ex)
            {
                //检测到异常，进行异常处理，防止应用程序崩溃
                MessageBox.Show("数据出现异常，请联系DBA检查数据库！\n原因：" + ex.Message);
            }
            finally
            {
                //关闭数据库连接，释放资源
                  //关闭
             if (conn != null)
                conn.Close();
            }
        }

        //单击停止按钮，停止计时器
        private void button2_Click(object sender, EventArgs e)
        {
            this.timer1.Stop();
        }

        //单击保存按钮，保存结果
        private void button3_Click(object sender, EventArgs e)
        {
            Save();
        }

        //保存方法
        private void Save()
        {
            //获取链接的同时打开链接
            SqlConnection conn = new SqlConnection("Data Source=ACER_ZHOU;Initial Catalog=Praise;Persist Security Info=True;User ID=sa;Password=123456");
            conn.Open();

            //创建sql语句

            string line = string.Format("INSERT INTO [Win] ([Name],[Department],[EmployeeID],[PrizeName],[Type]) VALUES('{0}','{1}','{2}','{3}','{4}')", this.textBox2.Text, this.textBox3.Text, this.textBox1.Text, label1.Text, label2.Text);

            //创建执行命令
            SqlCommand a1 = new SqlCommand(line, conn);
            int result = a1.ExecuteNonQuery();
            //关闭资源
            conn.Close();

            //根据返回受影响的行数判断数据是否保存成功
            if (result > 0)
            {   
                //奖品数量
                int i = int.Parse(label7.Text);
                i--;
                MessageBox.Show("保存成功");
                if (i == 2)
                {
                    this.label7.Text = "2";
                }
                if (i == 1 && this.label2.Text == "三等奖")
                {
                    this.label7.Text = "1";
                }
              
                //摇奖结束后显示摇奖结果
                if (i == 0 && this.label2.Text == "三等奖")
                {
                    WinForm form = new WinForm();
                    this.Hide();
                    form.Show();
                }
                if (i == 0 && this.label2.Text == "二等奖")
                {
                    this.label1.Text = "诺基亚手机";
                    this.label2.Text = "三等奖";
                    this.label7.Text = "3";
                }
               
               
                if (i == 0 && this.label2.Text == "一等奖")
                {
                    this.label1.Text = "苹果电脑";
                    this.label2.Text = "二等奖";
                    this.label7.Text = "2";
                }
                if (i == 1 && this.label2.Text == "二等奖")
                {
                    this.label7.Text = "1";
                }
                if (i == 0 && this.label2.Text == "特等奖")
                {
                    this.label1.Text = "威尼斯七日游";
                    this.label2.Text = "一等奖";
                    this.label7.Text = "2";
                }
                if (i == 1 && this.label2.Text == "一等奖")
                {
                    this.label7.Text = "1";
                }
            }
            else
            {
                MessageBox.Show("保存失败");
            }
        }

        }
	}

