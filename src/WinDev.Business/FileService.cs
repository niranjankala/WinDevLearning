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


        public List<Files> GetDirectoryHierarchy(string directoryPath, Func<Files, bool> exclusionFilter)
        {
            List<Files> files = new List<Files>();
            files.Add(new Files(Path.GetFileName(directoryPath), 0, true));
            files.AddRange(GetDirectoryFiles(directoryPath, 1, exclusionFilter));
            return files;

        }

        public void ExportDirectoryHierarchyToCSV(string directoryPath, string exportFile, Func<Files, bool> exclusionFilter)
        {
            using (StreamWriter tr = new StreamWriter(exportFile))
            {
                foreach (var filese in GetDirectoryHierarchy(directoryPath, exclusionFilter))
                {
                    tr.WriteLine(filese.ToString());
                }
                tr.Close();
            }
        }
        public List<Files> GetDirectoryFiles(string path, int i, Func<Files, bool> exclusionFilter)
        {
            List<Files> files = new List<Files>();
            DirectoryInfo directory = new DirectoryInfo(path);

            foreach (var d in directory.GetDirectories())
            {
                Files directoryFile = new Files(d.Name, i, true);
                if (!new List<Files>() { directoryFile }.Any(exclusionFilter))
                {
                    files.Add(directoryFile);
                    files.AddRange(GetDirectoryFiles(Path.Combine(path, d.Name), i + 1, exclusionFilter));
                }
            }
            foreach (var f in directory.GetFiles("*.cs"))
            {
                files.Add(new Files(f.Name, i));
            }
            return files;
        }
    }
}
