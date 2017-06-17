using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SwitchPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.pictureBox1.Show();
            try
            {

                String path = Directory.GetCurrentDirectory() + "\\word\\";

                //第一种方法
                var files = Directory.GetFiles(path, "*.doc");
                if (files != null && files.Count() > 0)
                {
                    List<ApplyInfo> lst = new List<ApplyInfo>();
                    SubStr s = new SubStr();
                    foreach (var file in files)
                    {
                        if (file.IndexOf('~') < 0)
                        {
                            string content = s.Doc2Text(file);
                            lst.Add(s.GetInfo(content));
                        }

                    }
                    //删除生成的文件
                    var files_xls = Directory.GetFiles(path, "*.xls");
                    foreach (var fp in files_xls)
                    {
                        System.IO.FileInfo file = new System.IO.FileInfo(fp);
                        if (file.Exists)//文件是否存在，存在则执行删除  
                        {
                            file.Delete();
                        }

                    }
                    //生成新导出的文件
                    DataTable dt = s.ListToDataTable(lst);
                    this.dataGridView1.DataSource = dt;
                   // CreateExcel_New(dt, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls");
                    MessageBox.Show("生成成功！");
                }
                else
                {
                    MessageBox.Show("未找到要转换的文件！");
                }


            }
            catch (Exception ex)
            {
                this.msg.Text = "工具异常，请联系开发人员：刘 18016227549";
            }
            this.pictureBox1.Hide();

        }

        /// <summary>
        /// 输出Excel
        /// </summary>
        /// <param name="exportFileName">名字</param>
        /// <param name="btfile"></param>
        public void CreateExcel_New(DataTable tb, string exportFileName)
        {
            string direct = Directory.GetCurrentDirectory() + "\\word\\";
            StreamWriter sw = new StreamWriter(direct + exportFileName, false, Encoding.GetEncoding("gb2312"));
            StringBuilder st = new StringBuilder();
            foreach (DataColumn dc in tb.Columns)
            {
                st.Append(dc.ColumnName + "\t");
            }
            foreach (DataRow dr in tb.Rows)
            {
                st.Append("\r");
                for (int i = 0; i < tb.Columns.Count; i++)
                {
                    if (dr[i] != null)
                    {
                        st.Append(dr[i].ToString().Replace("\t", "").Trim() + "\t");
                    }
                }
            }
            sw.Write(st.ToString());
            sw.Flush();
            sw.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string direct = Directory.GetCurrentDirectory() + "\\word\\";
            if (!Directory.Exists(direct))//若文件夹不存在则新建文件夹   
            {
                Directory.CreateDirectory(direct); //新建文件夹   
            }

            System.Diagnostics.Process.Start("Explorer.exe", direct);

        }
    }
}