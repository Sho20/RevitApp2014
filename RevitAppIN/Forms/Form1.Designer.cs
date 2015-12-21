namespace RevitAppIN.Forms
{
    partial class RevitAppMainForm
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
            this.testTxtbox = new System.Windows.Forms.TextBox();
            this.testBtn = new System.Windows.Forms.Button();
            this.testListBox = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // testTxtbox
            // 
            this.testTxtbox.Location = new System.Drawing.Point(12, 29);
            this.testTxtbox.Name = "testTxtbox";
            this.testTxtbox.Size = new System.Drawing.Size(101, 20);
            this.testTxtbox.TabIndex = 2;
            this.testTxtbox.Text = "600";
            // 
            // testBtn
            // 
            this.testBtn.Location = new System.Drawing.Point(12, 74);
            this.testBtn.Name = "testBtn";
            this.testBtn.Size = new System.Drawing.Size(101, 33);
            this.testBtn.TabIndex = 3;
            this.testBtn.Text = "Execute";
            this.testBtn.UseVisualStyleBackColor = true;
            this.testBtn.Click += new System.EventHandler(this.testBtn_Click);
            // 
            // testListBox
            // 
            this.testListBox.FormattingEnabled = true;
            this.testListBox.Location = new System.Drawing.Point(119, 12);
            this.testListBox.Name = "testListBox";
            this.testListBox.Size = new System.Drawing.Size(318, 238);
            this.testListBox.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 130);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(101, 39);
            this.button1.TabIndex = 5;
            this.button1.Text = "Update Connector Table";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // RevitAppMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(448, 262);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.testListBox);
            this.Controls.Add(this.testBtn);
            this.Controls.Add(this.testTxtbox);
            this.Name = "RevitAppMainForm";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox testTxtbox;
        private System.Windows.Forms.Button testBtn;
        private System.Windows.Forms.ListBox testListBox;
        private System.Windows.Forms.Button button1;
    }
}