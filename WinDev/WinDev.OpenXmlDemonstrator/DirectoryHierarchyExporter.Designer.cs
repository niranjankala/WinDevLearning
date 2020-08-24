namespace WinDev.OpenXmlDemonstrator
{
    partial class Form1
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
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.pathLbl = new System.Windows.Forms.Label();
            this.btnWriteHierarchyToExcel = new System.Windows.Forms.Button();
            this.txtDirectoryPath = new System.Windows.Forms.TextBox();
            this.btnGetFiles = new System.Windows.Forms.Button();
            this.btnBrowseDocExportP1 = new System.Windows.Forms.Button();
            this.lblExportUrlP1 = new System.Windows.Forms.Label();
            this.txtExportPath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // pathLbl
            // 
            this.pathLbl.AutoSize = true;
            this.pathLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pathLbl.Location = new System.Drawing.Point(12, 9);
            this.pathLbl.Name = "pathLbl";
            this.pathLbl.Size = new System.Drawing.Size(225, 13);
            this.pathLbl.TabIndex = 4;
            this.pathLbl.Text = "Energy Plus  source code folder path :";
            // 
            // btnWriteHierarchyToExcel
            // 
            this.btnWriteHierarchyToExcel.Location = new System.Drawing.Point(146, 105);
            this.btnWriteHierarchyToExcel.Name = "btnWriteHierarchyToExcel";
            this.btnWriteHierarchyToExcel.Size = new System.Drawing.Size(224, 24);
            this.btnWriteHierarchyToExcel.TabIndex = 7;
            this.btnWriteHierarchyToExcel.Text = "Write Hierarchy To Excel";
            this.btnWriteHierarchyToExcel.UseVisualStyleBackColor = true;
            this.btnWriteHierarchyToExcel.Click += new System.EventHandler(this.btnWriteHierarchyToExcel_Click);
            // 
            // txtDirectoryPath
            // 
            this.txtDirectoryPath.Location = new System.Drawing.Point(15, 25);
            this.txtDirectoryPath.Name = "txtDirectoryPath";
            this.txtDirectoryPath.Size = new System.Drawing.Size(407, 20);
            this.txtDirectoryPath.TabIndex = 5;
            // 
            // btnGetFiles
            // 
            this.btnGetFiles.Location = new System.Drawing.Point(434, 23);
            this.btnGetFiles.Name = "btnGetFiles";
            this.btnGetFiles.Size = new System.Drawing.Size(75, 23);
            this.btnGetFiles.TabIndex = 6;
            this.btnGetFiles.Text = "Browse";
            this.btnGetFiles.UseVisualStyleBackColor = true;
            this.btnGetFiles.Click += new System.EventHandler(this.btnGetFiles_Click);
            // 
            // btnBrowseDocExportP1
            // 
            this.btnBrowseDocExportP1.Location = new System.Drawing.Point(435, 61);
            this.btnBrowseDocExportP1.Name = "btnBrowseDocExportP1";
            this.btnBrowseDocExportP1.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseDocExportP1.TabIndex = 13;
            this.btnBrowseDocExportP1.Text = "Browse";
            this.btnBrowseDocExportP1.UseVisualStyleBackColor = true;
            this.btnBrowseDocExportP1.Click += new System.EventHandler(this.btnBrowseDocExportP1_Click);
            // 
            // lblExportUrlP1
            // 
            this.lblExportUrlP1.AutoSize = true;
            this.lblExportUrlP1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblExportUrlP1.Location = new System.Drawing.Point(12, 48);
            this.lblExportUrlP1.Name = "lblExportUrlP1";
            this.lblExportUrlP1.Size = new System.Drawing.Size(144, 13);
            this.lblExportUrlP1.TabIndex = 11;
            this.lblExportUrlP1.Text = "Excel Export File Path  :";
            // 
            // txtExportPath
            // 
            this.txtExportPath.Location = new System.Drawing.Point(13, 64);
            this.txtExportPath.Name = "txtExportPath";
            this.txtExportPath.Size = new System.Drawing.Size(409, 20);
            this.txtExportPath.TabIndex = 12;
            this.txtExportPath.Text = "D:\\";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(522, 141);
            this.Controls.Add(this.btnBrowseDocExportP1);
            this.Controls.Add(this.lblExportUrlP1);
            this.Controls.Add(this.txtExportPath);
            this.Controls.Add(this.pathLbl);
            this.Controls.Add(this.btnWriteHierarchyToExcel);
            this.Controls.Add(this.txtDirectoryPath);
            this.Controls.Add(this.btnGetFiles);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label pathLbl;
        private System.Windows.Forms.Button btnWriteHierarchyToExcel;
        private System.Windows.Forms.TextBox txtDirectoryPath;
        private System.Windows.Forms.Button btnGetFiles;
        private System.Windows.Forms.Button btnBrowseDocExportP1;
        private System.Windows.Forms.Label lblExportUrlP1;
        private System.Windows.Forms.TextBox txtExportPath;
    }
}

