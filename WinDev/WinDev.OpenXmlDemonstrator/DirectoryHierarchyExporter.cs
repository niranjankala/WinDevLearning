using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinDev.Business.OpenXml;

namespace WinDev.OpenXmlDemonstrator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGetFiles_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderPath = new FolderBrowserDialog();
            System.Windows.Forms.DialogResult dr = folderPath.ShowDialog();
            if (dr.Equals(DialogResult.OK))
            {
                txtDirectoryPath.Text = folderPath.SelectedPath;

            }
        }

        private void btnWriteHierarchyToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelManager excelManager = new ExcelManager();
                //method to read energy plus error message from files and generate excel file for the error messages
                excelManager.ExportDirectoryHierarchy(txtDirectoryPath.Text, txtExportPath.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnBrowseDocExportP1_Click(object sender, EventArgs e)
        {
            SaveFileDialog filePath = new SaveFileDialog();
            System.Windows.Forms.DialogResult dr = filePath.ShowDialog();
            if (dr.Equals(DialogResult.OK))
            {
                txtExportPath.Text = filePath.FileName;
            }
        }
    }
}
