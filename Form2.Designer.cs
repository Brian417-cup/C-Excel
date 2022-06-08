namespace CSharpForExcel
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SQLResultShowDataGridView = new System.Windows.Forms.DataGridView();
            this.exportToExcelBtn = new System.Windows.Forms.Button();
            this.exportToCsvBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.SQLResultShowDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // SQLResultShowDataGridView
            // 
            this.SQLResultShowDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.SQLResultShowDataGridView.Location = new System.Drawing.Point(39, 44);
            this.SQLResultShowDataGridView.Name = "SQLResultShowDataGridView";
            this.SQLResultShowDataGridView.RowTemplate.Height = 23;
            this.SQLResultShowDataGridView.Size = new System.Drawing.Size(516, 332);
            this.SQLResultShowDataGridView.TabIndex = 0;
            // 
            // exportToExcelBtn
            // 
            this.exportToExcelBtn.Location = new System.Drawing.Point(593, 111);
            this.exportToExcelBtn.Name = "exportToExcelBtn";
            this.exportToExcelBtn.Size = new System.Drawing.Size(107, 63);
            this.exportToExcelBtn.TabIndex = 1;
            this.exportToExcelBtn.Text = "将查询结果导出为Office Excel";
            this.exportToExcelBtn.UseVisualStyleBackColor = true;
            this.exportToExcelBtn.Click += new System.EventHandler(this.exportToExcelBtn_Click);
            // 
            // exportToCsvBtn
            // 
            this.exportToCsvBtn.Location = new System.Drawing.Point(593, 244);
            this.exportToCsvBtn.Name = "exportToCsvBtn";
            this.exportToCsvBtn.Size = new System.Drawing.Size(107, 63);
            this.exportToCsvBtn.TabIndex = 2;
            this.exportToCsvBtn.Text = "将查询结果导出为Csv";
            this.exportToCsvBtn.UseVisualStyleBackColor = true;
            this.exportToCsvBtn.Click += new System.EventHandler(this.exportToCsvBtn_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(740, 437);
            this.Controls.Add(this.exportToCsvBtn);
            this.Controls.Add(this.exportToExcelBtn);
            this.Controls.Add(this.SQLResultShowDataGridView);
            this.Name = "Form2";
            this.Text = "查询结果";
            ((System.ComponentModel.ISupportInitialize)(this.SQLResultShowDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView SQLResultShowDataGridView;
        private System.Windows.Forms.Button exportToExcelBtn;
        private System.Windows.Forms.Button exportToCsvBtn;
    }
}