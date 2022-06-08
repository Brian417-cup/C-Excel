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
    public partial class Form2 : Form
    {
        DataTable result = null;

        public Form2()
        {
            InitializeComponent();
        }

        public Form2(DataTable src)
        {
            InitializeComponent();
            SQLResultShowDataGridView.DataSource = src;
            MessageBox.Show("SQL查询语句执行成功");
        }

        private void exportToExcelBtn_Click(object sender, EventArgs e)
        {
            //跳出显示对话框
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = "导出的Excel名字";
            saveFileName = saveDialog.FileName;
            DialogResult result = saveDialog.ShowDialog();
            if (result == DialogResult.Cancel) return; //被点了取消
            else if (result == DialogResult.OK)
            {
                //MessageBox.Show(saveDialog.FileName);

                //选择保存
                SimpleExcelTool simpleExcelTool = new SimpleExcelTool();
                if (
                    simpleExcelTool.ExportExcels(saveDialog.FileName, SQLResultShowDataGridView)
                    )
                {
                    MessageBox.Show("保存成功,路径为:" + saveDialog.FileName);
                }
                else
                {
                    MessageBox.Show("保存失败");
                }
            }
        }

        private void exportToCsvBtn_Click(object sender, EventArgs e)
        {
            //跳出显示对话框
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "csv";
            saveDialog.Filter = "CSV 文件 (*.csv)|*.csv";
            saveDialog.FileName = "导出的Csv名字";
            saveFileName = saveDialog.FileName;
            DialogResult result = saveDialog.ShowDialog();
            if (result == DialogResult.Cancel) return; //被点了取消
            else if (result == DialogResult.OK)
            {
                //MessageBox.Show(saveDialog.FileName);

                //选择保存
                SimpleExcelTool simpleExcelTool = new SimpleExcelTool();
                if (
                    simpleExcelTool.ExportCSV(saveDialog.OpenFile(), saveDialog.FileName, SQLResultShowDataGridView)
                    )
                {
                    MessageBox.Show("保存成功,路径为:" + saveDialog.FileName);
                }
                else
                {
                    MessageBox.Show("保存失败");
                }
            }
        }


    }
}
