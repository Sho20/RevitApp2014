namespace RevitAppIN.Forms
{
    partial class ReportForm
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
            this.RunReport = new System.Windows.Forms.Button();
            this.txtInputFilePath = new System.Windows.Forms.TextBox();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSetFilePath = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // RunReport
            // 
            this.RunReport.Location = new System.Drawing.Point(329, 68);
            this.RunReport.Name = "RunReport";
            this.RunReport.Size = new System.Drawing.Size(118, 27);
            this.RunReport.TabIndex = 0;
            this.RunReport.Text = "Run Excel Report";
            this.RunReport.UseVisualStyleBackColor = true;
            this.RunReport.Click += new System.EventHandler(this.RunReport_Click);
            // 
            // txtInputFilePath
            // 
            this.txtInputFilePath.Location = new System.Drawing.Point(102, 27);
            this.txtInputFilePath.Name = "txtInputFilePath";
            this.txtInputFilePath.Size = new System.Drawing.Size(207, 20);
            this.txtInputFilePath.TabIndex = 24;
            this.txtInputFilePath.Text = "txtInputFilePath";
            this.txtInputFilePath.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.txtInputFilePath_MouseDoubleClick);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "xlsx";
            this.saveFileDialog1.Filter = "Excel xlsx|*.xlsx";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 25;
            this.label1.Text = "Export Report to";
            // 
            // btnSetFilePath
            // 
            this.btnSetFilePath.Location = new System.Drawing.Point(329, 23);
            this.btnSetFilePath.Name = "btnSetFilePath";
            this.btnSetFilePath.Size = new System.Drawing.Size(118, 27);
            this.btnSetFilePath.TabIndex = 26;
            this.btnSetFilePath.Text = "Set File Path";
            this.btnSetFilePath.UseVisualStyleBackColor = true;
            this.btnSetFilePath.Click += new System.EventHandler(this.btnSetFilePath_Click);
            // 
            // ReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(480, 141);
            this.Controls.Add(this.btnSetFilePath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtInputFilePath);
            this.Controls.Add(this.RunReport);
            this.Name = "ReportForm";
            this.Text = "ReportForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button RunReport;
        private System.Windows.Forms.TextBox txtInputFilePath;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSetFilePath;
    }
}