namespace ExcelEditor
{
    partial class FileSelector
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
            this.openExcelFile = new System.Windows.Forms.OpenFileDialog();
            this.openFileButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openExcelFile
            // 
            this.openExcelFile.DefaultExt = "*.xls";
            this.openExcelFile.FileName = "openExcelFile";
            this.openExcelFile.FileOk += new System.ComponentModel.CancelEventHandler(this.openExcelFile_FileOk);
            this.openExcelFile.HelpRequest += new System.EventHandler(this.openExcelFile_HelpRequest);
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(225, 65);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(122, 60);
            this.openFileButton.TabIndex = 0;
            this.openFileButton.Text = "Open Menu";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.openFileButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(185, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(194, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Find the excel file by clikcing this button";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // FileSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(565, 191);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.openFileButton);
            this.MaximizeBox = false;
            this.Name = "FileSelector";
            this.Text = "Excel Editor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openExcelFile;
        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.Label label1;

    }
}

