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
using Aspose.Cells;

namespace yaojiang
{
	public partial class StartForm : Form
	{
        private string openFileName = null;
		public StartForm()
		{
			InitializeComponent();
		}
        //点击启动程序按钮事件
		private void button1_Click(object sender, EventArgs e)
		{
            //打开摇奖窗体并隐藏当前窗体
		    LuckyForm form = new LuckyForm();
		    form.Show();
            this.Hide();
		}

        //单击浏览按钮事件
		private void button2_Click(object sender, EventArgs e)
		{
		    this.dlgOpen_XLSX.InitialDirectory = Application.StartupPath;
			DialogResult r = this.dlgOpen_XLSX.ShowDialog();
			if(r == DialogResult.OK)
			{
				openFileName = this.dlgOpen_XLSX.FileName;
				// 实例化 Workbook，在其构造函数中传入要打开的文件名即可打开文件
				Workbook wb = new Workbook(openFileName);
				// 获取 下标为 0 的工作表的名称
				string workSheetName = wb.Worksheets[0].Name;
				// 在窗体标题上显示该名称
				this.Text = workSheetName;
				// 获取下标为 0 的工作表的所有单元格
				Cells cells = wb.Worksheets[0].Cells;
                for (int row = 1; row < cells.MaxDataRow + 1; row++)
                {
                    // 读取第 row 行的下标为 1 的列（下标从 0 开始）
                    string id = cells[row, 1].StringValue.Trim();	// 读取员工编号
                    // 读取第 row 行的下标为 2 的列（下标从 0 开始）
                    string name = cells[row, 2].StringValue.Trim();	// 读取员工姓名
                    //读取第 row 行的下标为 3 的列（下标从 0 开始）
                    string dep = cells[row, 3].StringValue.Trim();  // 读取员工所在部门
                }
		}
            //在textbox中显示文件路径
			this.textBox1.Text = this.dlgOpen_XLSX.FileName;

            //创建连接对象
            string connstr = "Data Source=.;Initial Catalog=Praise;Persist Security Info=True;User ID=sa;Password=123456";
            SqlConnection conn = new SqlConnection(connstr);
            conn.Open();
            //创建sql语句
            string sql = "select count(EmployeeID) from Employee";
            SqlCommand cmd = new SqlCommand(sql, conn);
            int count = (int)cmd.ExecuteScalar();
            this.textBox2.Text = count.ToString();
            conn.Close();
        }

        //单击浏览按钮事件
		private void button4_Click(object sender, EventArgs e)
		{
		    this.dlgOpen_XLSX.InitialDirectory = Application.StartupPath;
			DialogResult r = this.dlgOpen_XLSX.ShowDialog();
			if(r == DialogResult.OK)
			{
				openFileName = this.dlgOpen_XLSX.FileName;
				// 实例化 Workbook，在其构造函数中传入要打开的文件名即可打开文件
				Workbook wb = new Workbook(openFileName);
				// 获取 下标为 0 的工作表的名称
				string workSheetName = wb.Worksheets[0].Name;
                //// 在窗体标题上显示该名称
                //this.Text = workSheetName;
				// 获取下标为 0 的工作表的所有单元格
				Cells cells = wb.Worksheets[0].Cells;
				for(int row = 1; row < cells.MaxDataRow + 1; row++)
				{
					// 读取第 row 行的下标为 1 的列（下标从 0 开始）
					string priid = cells[row, 1].StringValue.Trim();	// 读取员工编号
					// 读取第 row 行的下标为 2 的列（下标从 0 开始）
					string priname = cells[row, 2].StringValue.Trim();	// 读取员工姓名
					//读取第 row 行的下标为 3 的列（下标从 0 开始）
					string type = cells[row, 3].StringValue.Trim();  // 读取员工所在部门
					//读取第 row 行的下标为 4 的列（下标从 0 开始）
					string num = cells[row, 4].StringValue.Trim();
				}
                //在textbox中显示文件路径
            this.textBox3.Text = this.dlgOpen_XLSX.FileName;

            //创建连接对象
            string connstr = "Data Source=.;Initial Catalog=Praise;Persist Security Info=True;User ID=sa;Password=123456";
            SqlConnection con = new SqlConnection(connstr);
            con.Open();
            //创建sql语句
            string sql1 = "select SUM(Number) from Prize";
            SqlCommand cmd1 = new SqlCommand(sql1, con);
            int count = (int)cmd1.ExecuteScalar();
            this.textBox4.Text = count.ToString();
            con.Close();
			}
		}
	  
	}
}
