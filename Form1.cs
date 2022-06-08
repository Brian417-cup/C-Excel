using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CSharpForExcel
{
    public partial class Form1 : Form
    {
        private string excelPath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void openExcelBtn_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelPath = openFileDialog.FileName;
                    excelPathShowArea.Text = excelPath;
                    //MessageBox.Show(openFileDialog.FileName);
                }
                else
                {
                    MessageBox.Show("文件打开失败!!");

                }
            }
        }

        private void SQLQueryOpBtn_Click(object sender, EventArgs e)
        {
            string sqlData = SQLInputTxt.Text;
            if (sqlData.Length == 0)
            {
                MessageBox.Show("输入不能为空");
                return;
            }

            try
            {
                SimpleExcelTool simpleExcelTool = new SimpleExcelTool();
                simpleExcelTool.openConnection(excelPath);
                simpleExcelTool.clearLastQueryResult();
                DataSet queryResult = simpleExcelTool.executeExcelQuerySQL(sqlData);
                new Form2(queryResult.Tables[0]).Show();
                simpleExcelTool.closeConnection();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
            }
        }

        private void SQLInsertAndDeleteOpBtn_Click(object sender, EventArgs e)
        {
            string sqlData = SQLInputTxt.Text;
            if (sqlData.Length == 0)
            {
                MessageBox.Show("输入不能为空");
                return;
            }

            try
            {
                SimpleExcelTool simpleExcelTool = new SimpleExcelTool();
                simpleExcelTool.openConnection(excelPath);
                simpleExcelTool.clearLastQueryResult();
                int influencedRowsCnt = simpleExcelTool.executeExcelInsertAndDeleteSQL(sqlData);
                simpleExcelTool.closeConnection();

                if (influencedRowsCnt != -1)
                {
                    MessageBox.Show("执行成功，一共有" + influencedRowsCnt.ToString() + "行受影响");
                }

            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
            }
        }
    }
}
