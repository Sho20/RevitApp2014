namespace RevitAppIN.Forms
{
    partial class ImportForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtInputFilePath = new System.Windows.Forms.TextBox();
            this.btnSetFilePath = new System.Windows.Forms.Button();
            this.RunReport = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 26;
            this.label1.Text = "Import Report from";
            // 
            // txtInputFilePath
            // 
            this.txtInputFilePath.Location = new System.Drawing.Point(114, 27);
            this.txtInputFilePath.Name = "txtInputFilePath";
            this.txtInputFilePath.Size = new System.Drawing.Size(207, 20);
            this.txtInputFilePath.TabIndex = 27;
            // 
            // btnSetFilePath
            // 
            this.btnSetFilePath.Location = new System.Drawing.Point(339, 23);
            this.btnSetFilePath.Name = "btnSetFilePath";
            this.btnSetFilePath.Size = new System.Drawing.Size(118, 27);
            this.btnSetFilePath.TabIndex = 28;
            this.btnSetFilePath.Text = "Set File Path";
            this.btnSetFilePath.UseVisualStyleBackColor = true;
            // 
            // RunReport
            // 
            this.RunReport.Location = new System.Drawing.Point(339, 68);
            this.RunReport.Name = "RunReport";
            this.RunReport.Size = new System.Drawing.Size(118, 27);
            this.RunReport.TabIndex = 29;
            this.RunReport.Text = "Import Report";
            this.RunReport.UseVisualStyleBackColor = true;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Excel xlsx|*.xlsx";
            // 
            // ImportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 142);
            this.Controls.Add(this.RunReport);
            this.Controls.Add(this.btnSetFilePath);
            this.Controls.Add(this.txtInputFilePath);
            this.Controls.Add(this.label1);
            this.Name = "ImportForm";
            this.Text = "ImportForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtInputFilePath;
        private System.Windows.Forms.Button btnSetFilePath;
        private System.Windows.Forms.Button RunReport;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}