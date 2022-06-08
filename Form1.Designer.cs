namespace CSharpForExcel
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.openExcelBtn = new System.Windows.Forms.Button();
            this.excelPathShowArea = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.SQLInsertAndDeleteOpBtn = new System.Windows.Forms.Button();
            this.SQLQueryOpBtn = new System.Windows.Forms.Button();
            this.SQLInputTxt = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openExcelBtn
            // 
            this.openExcelBtn.Location = new System.Drawing.Point(27, 42);
            this.openExcelBtn.Name = "openExcelBtn";
            this.openExcelBtn.Size = new System.Drawing.Size(116, 36);
            this.openExcelBtn.TabIndex = 1;
            this.openExcelBtn.Text = "选择Excel文件";
            this.openExcelBtn.UseVisualStyleBackColor = true;
            this.openExcelBtn.Click += new System.EventHandler(this.openExcelBtn_Click);
            // 
            // excelPathShowArea
            // 
            this.excelPathShowArea.Location = new System.Drawing.Point(166, 51);
            this.excelPathShowArea.Name = "excelPathShowArea";
            this.excelPathShowArea.Size = new System.Drawing.Size(478, 21);
            this.excelPathShowArea.TabIndex = 3;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.SQLInsertAndDeleteOpBtn);
            this.panel1.Controls.Add(this.SQLQueryOpBtn);
            this.panel1.Controls.Add(this.SQLInputTxt);
            this.panel1.Location = new System.Drawing.Point(27, 113);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(617, 262);
            this.panel1.TabIndex = 4;
            // 
            // SQLInsertAndDeleteOpBtn
            // 
            this.SQLInsertAndDeleteOpBtn.Location = new System.Drawing.Point(420, 226);
            this.SQLInsertAndDeleteOpBtn.Name = "SQLInsertAndDeleteOpBtn";
            this.SQLInsertAndDeleteOpBtn.Size = new System.Drawing.Size(144, 33);
            this.SQLInsertAndDeleteOpBtn.TabIndex = 3;
            this.SQLInsertAndDeleteOpBtn.Text = "执行SQL更新和插入语句";
            this.SQLInsertAndDeleteOpBtn.UseVisualStyleBackColor = true;
            this.SQLInsertAndDeleteOpBtn.Click += new System.EventHandler(this.SQLInsertAndDeleteOpBtn_Click);
            // 
            // SQLQueryOpBtn
            // 
            this.SQLQueryOpBtn.Location = new System.Drawing.Point(49, 226);
            this.SQLQueryOpBtn.Name = "SQLQueryOpBtn";
            this.SQLQueryOpBtn.Size = new System.Drawing.Size(144, 33);
            this.SQLQueryOpBtn.TabIndex = 2;
            this.SQLQueryOpBtn.Text = "执行SQL查询语句";
            this.SQLQueryOpBtn.UseVisualStyleBackColor = true;
            this.SQLQueryOpBtn.Click += new System.EventHandler(this.SQLQueryOpBtn_Click);
            // 
            // SQLInputTxt
            // 
            this.SQLInputTxt.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SQLInputTxt.Location = new System.Drawing.Point(18, 57);
            this.SQLInputTxt.Multiline = true;
            this.SQLInputTxt.Name = "SQLInputTxt";
            this.SQLInputTxt.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.SQLInputTxt.Size = new System.Drawing.Size(574, 163);
            this.SQLInputTxt.TabIndex = 1;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(18, 3);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(596, 36);
            this.textBox1.TabIndex = 4;
            this.textBox1.Text = "输入SQL语句 例:SELECT * FROM [Sheet1$] \\ insert into [Sheet1$A32:D32] values(123,123,1" +
                "23,123) \\ update from [Sheet1$] set ... where ...";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(674, 416);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.excelPathShowArea);
            this.Controls.Add(this.openExcelBtn);
            this.Name = "Form1";
            this.Text = "ExcekSQLSimpleTools";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openExcelBtn;
        private System.Windows.Forms.TextBox excelPathShowArea;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox SQLInputTxt;
        private System.Windows.Forms.Button SQLQueryOpBtn;
        private System.Windows.Forms.Button SQLInsertAndDeleteOpBtn;
        private System.Windows.Forms.TextBox textBox1;
    }
}

