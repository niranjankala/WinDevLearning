using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinDev.StorageExplorer
{
    class FileSystemEntry
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsDirectory { get; set; }
        public List<FileSystemEntry> Files { get; set; }
        public string URL { get; set; }
        public DateTime CreationDate { get; set; }
        public DateTime LastModified { get; set; }
    }
}
