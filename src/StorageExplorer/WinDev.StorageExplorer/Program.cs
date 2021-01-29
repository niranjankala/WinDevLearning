using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinDev.StorageExplorer
{
    class Program
    {
        static string hostURL = string.Empty;
        static void Main(string[] args)
        {
            hostURL = ConfigurationManager.AppSettings["WebApplicationURL"];
            //List<FileSystemEntry> samples = CreateSamplesInfo(Application.StartupPath);
            //GenerateXMLFile(samples, $"{Application.StartupPath}\\StorageFilesInfo.xml");
            //GenerateHtmlFile(samples, $"{Application.StartupPath}\\Index.html");
            CreateZipFile();
            Console.WriteLine("File created...");
            Console.ReadKey();
        }

        static void CreateZipFile()
        {
            string directoryPath = @"C:\Users\niranjansingh\Desktop\R&R";

            using (var fs = new FileStream($@"C:\Users\niranjansingh\Desktop\1.zip", FileMode.Create))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                zip.CreateEntryFromAny(directoryPath); // just end with "/"
            }
        }
        private static List<FileSystemEntry> CreateSamplesInfo(string dirPath)
        {
            List<FileSystemEntry> sampleList = new List<FileSystemEntry>();

            if (Directory.Exists(dirPath))
            {
                string webhostDirPath = string.Empty;
                DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
                string sampleFileVersion = directoryInfo.Name;                
                string[] childDirectories = Directory.GetDirectories(dirPath);
                foreach (string d in childDirectories)
                {
                    sampleList.Add(GetDirectoryInfo(d, sampleFileVersion));
                }
                //sampleList.AddRange(GetFilesInfo(directoryInfo.GetFiles("*.zip"), sampleFileVersion));
            }
            return sampleList;

        }

        private static FileSystemEntry GetDirectoryInfo(string dirPath, string sampleFileVersion)
        {
            FileSystemEntry FileSystemEntry = null;
            try
            { 
                if (Directory.Exists(dirPath))
                {
                    //Directory.CreateDirectory(dirPath);

                    FileSystemEntry = new FileSystemEntry();
                    DirectoryInfo sampleFolderInfo = new DirectoryInfo(dirPath);
                    FileSystemEntry.FileName = sampleFolderInfo.Name;
                    FileSystemEntry.FilePath = sampleFolderInfo.FullName.Replace($"{Application.StartupPath}\\", "");
                    FileSystemEntry.IsDirectory = true;
                    FileInfo[] files = sampleFolderInfo.GetFiles();
                    string[] childDirectories = Directory.GetDirectories(dirPath);
                    if (files.Length > 0 || childDirectories.Length > 0)
                    {
                        FileSystemEntry.Files = GetFilesInfo(files, sampleFileVersion);
                        foreach (string d in childDirectories)
                        {
                            FileSystemEntry.Files.Add(GetDirectoryInfo(d, sampleFileVersion));
                        }
                    }
                }
            }
            catch (System.Exception excpt)
            {
                MessageBox.Show(excpt.Message);
            }
            return FileSystemEntry;
        }

        private static List<FileSystemEntry> GetFilesInfo(FileInfo[] files, string sampleFileVersion)
        {
            List<FileSystemEntry> filesInfo = new List<FileSystemEntry>();
            foreach (FileInfo f in files)
            {
                filesInfo.Add(new FileSystemEntry()
                {
                    FileName = f.Name,
                    FilePath = f.FullName.Replace($"{Application.StartupPath}\\", ""),
                    URL = $"{hostURL}/{f.FullName.Replace($"{Application.StartupPath}\\", "").Replace("\\", "/")}",
                    CreationDate = f.CreationTimeUtc,
                    IsDirectory = false,
                    LastModified = f.LastWriteTimeUtc
                });
            }
            return filesInfo;
        }
        static void GenerateHtmlFile(List<FileSystemEntry> samples, string outputHtmlPath)
        {
            
        }

        private static void GenerateXMLFile(List<FileSystemEntry> samples, string outputXMLPath)
        {
            System.Xml.Serialization.XmlSerializer xs1 = new System.Xml.Serialization.XmlSerializer(typeof(List<FileSystemEntry>));
            System.IO.StreamWriter tw = new System.IO.StreamWriter(outputXMLPath, false);
            xs1.Serialize(tw, samples);
            tw.Close();
        }
    }
}
public static class ZipArchiveExtension
{

    public static void CreateEntryFromAny(this ZipArchive archive, string sourceName, string entryName = "")
    {
        var fileName = Path.GetFileName(sourceName);
        if (File.GetAttributes(sourceName).HasFlag(FileAttributes.Directory))
        {
            archive.CreateEntryFromDirectory(sourceName, Path.Combine(entryName, fileName));
        }
        else
        {
            archive.CreateEntryFromFile(sourceName, Path.Combine(entryName, fileName), CompressionLevel.Fastest);
            
        }
    }

    public static void CreateEntryFromDirectory(this ZipArchive archive, string sourceDirName, string entryName = "")
    {
        string[] files = Directory.GetFiles(sourceDirName).Concat(Directory.GetDirectories(sourceDirName)).ToArray();
        archive.CreateEntry(Path.Combine(entryName, Path.GetFileName(sourceDirName)));
        foreach (var file in files)
        {
            archive.CreateEntryFromAny(file, entryName);
        }
    }

    public static void CreateEntryFromFile(this ZipArchive archive, string sourceName, string entryName , CompressionLevel compressionLevel)
    {
        //var relativePath = filePath.Replace(InputDirectory, string.Empty);
        using (Stream fileStream = new FileStream(sourceName, FileMode.Open, FileAccess.Read))
        using (Stream fileStreamInZip = archive.CreateEntry(entryName).Open())
            fileStream.CopyTo(fileStreamInZip);

    }


}