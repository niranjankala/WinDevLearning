using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinDev.Common.DataObjects;

namespace WinDev.Business
{
    public class FileService
    {
        

        public List<Files> GetDirectoryHierarchy(string directoryPath)
        {
            List<Files> files = new List<Files>();
            directoryPath = @"D:\DevWorkSpaces\DA\Simergy_TFS\Simergy\Dev_3_3";

            files.Add(new Files(Path.GetFileName(directoryPath), 0));
            files.AddRange(GetDirectoryFiles(directoryPath, 1));           
            return files;

        }

        public void ExportDirectoryHierarchyToCSV(string directoryPath, string exportFile)
        {
            using (StreamWriter tr = new StreamWriter(exportFile))
            {
                foreach (var filese in GetDirectoryHierarchy(directoryPath))
                {
                    tr.WriteLine(filese.ToString());
                }
                tr.Close();
            }
        }
        public List<Files> GetDirectoryFiles(string path, int i)
        {
            List<Files> files = new List<Files>();
            DirectoryInfo directory = new DirectoryInfo(path);
            foreach (var d in directory.GetDirectories())
            {
                files.Add(new Files(d.Name, i));
                files.AddRange(GetDirectoryFiles(Path.Combine(path, d.Name), i + 1));
            }
            foreach (var f in directory.GetFiles("*.cs"))
            {
                files.Add(new Files(f.Name, i));
            }
            return files;
        }
    }
}
