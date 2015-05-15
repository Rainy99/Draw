using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace yaojiang
{
    public partial class WinForm : Form
    {
        public WinForm()
        {
            InitializeComponent();
        }
        DataTable dt;
        SqlDataAdapter dap;
        private void Form3_Load(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection("server=.;database = Praise;uid = sa;pwd = 123456;");
            string sql = "select * from Win";
            //创建datatable
            dt = new DataTable();
            //创建dataAdapter
            dap = new SqlDataAdapter(sql, conn);
            //使用fill();
            dap.Fill(dt);
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void 保存为Excel文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                //return false;
            }
            //创建Excel对象
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Application.Workbooks.Add(true);

            //生成字段名称
            for (int i = 1; i < dataGridView1.ColumnCount; i++)
            {
                excel.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
            }
            //填充数据
            for (int i = 1; i < dataGridView1.RowCount - 1; i++)   //循环行
            {
                for (int j = 1; j < dataGridView1.ColumnCount; j++) //循环列
                {
                    if (dataGridView1[j, i].ValueType == typeof(string))
                    {
                        excel.Cells[i + 2, j + 1] = "'" + dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                    else
                    {
                        excel.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            //设置禁止弹出保存和覆盖的询问提示框  
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.AlertBeforeOverwriting = false;

            //保存到临时工作簿
            //excel.Application.Workbooks.Add(true).Save();
            //保存文件

            excel.Save("D:" + "\\234.xls");
            excel.Quit();
            //return true;
        }

        //窗体关闭事件，清空数据库
        private void clear(object sender, FormClosingEventArgs e)
        {
            SqlConnection conn;
            string connStr = "server=.;database = Praise;uid = sa;pwd = 123456;";
            conn = new SqlConnection(connStr);
            conn.Open();

            string clear = "delete from Win";
            SqlCommand c = new SqlCommand(clear, conn);
            int r = c.ExecuteNonQuery();
            conn.Close();

            Application.Exit();

        }
    }
}
