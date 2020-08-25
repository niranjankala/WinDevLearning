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
using WinDev.Common.DataObjects;

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
                Func<Files, bool> exclusionFilter = (Files f) => (f.hasChildren && (f.name.ToLower() == "bin"
                || f.name.ToLower() == "obj"                
                || f.name.ToLower() == "properties"
                || f.name.ToLower() == ".vs"
                || f.name.ToLower() == "mymodules"
                || f.name.ToLower() == "prebuild_keepthisfirstinbuildorder"
                || f.name.ToLower() == "samplecode"
                || f.name.ToLower() == "thirdpartylibraries"
                || f.name.ToLower() == "packages"
                || f.name.ToLower() == "tests"
                || f.name.ToLower() == "templates"
                || f.name.ToLower() == "help"
                || f.name.ToLower() == "images"
                || f.name.ToLower() == "geomlibtests"
                ));
                excelManager.ExportDirectoryHierarchy(txtDirectoryPath.Text, txtExportPath.Text, exclusionFilter);
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
