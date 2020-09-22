using System;
using System.Collections.Generic;
using System.Data;
using System.Xml;
using System.IO;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace WinDev.Common.Helpers
{
    public class XMLHelper
    {
        private static string GetXMLData(string fileName)
        {
            FileStream file = null;
            StreamReader sr = null;
            try
            {
                file = new System.IO.FileStream(fileName, FileMode.Open, FileAccess.Read);
                sr = new System.IO.StreamReader(file);
                string xmlString = sr.ReadToEnd();
                sr.Close();
                file.Close();
                return xmlString;
            }
            finally
            {
                if ((file != null))
                {
                    file.Dispose();
                }
                if ((sr != null))
                {
                    sr.Dispose();
                }
            }
        }

        public static XmlDocument GetXMLDocumentFromFile(string fileName)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(GetXMLData(fileName));
            //doc.LoadXml(fileName);
            return doc;
        }

        public static XmlDocument GetXMLDocumentFromString(string xmlData)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlData);
            return doc;
        }

        public static IEnumerable<T> Convert<T>(System.Collections.IEnumerable enumerable)
        {
            foreach (object current in enumerable)
            {
                yield return (T)current;
            }
        }

        public static void CopyNodesToRoot(XmlDocument destinationDOM, XmlNodeList nodesToCopy)
        {
            XmlNode[] nodes;
            nodes = (new List<XmlNode>(XMLHelper.Convert<XmlNode>(nodesToCopy))).ToArray();
            for (int i = 0; i < nodes.Length; i++)
            {
                //Import node to main document and append to child
                destinationDOM.ChildNodes[0].AppendChild(destinationDOM.ImportNode(nodes[i], true));
            }

        }

        public static void CopyNode(XmlDocument destinationDOM, XmlNode nodeToCopy)
        {
            //Import node to main document and append to child
            XmlNode nodeToMove = destinationDOM.ImportNode(nodeToCopy, true);
            destinationDOM.AppendChild(nodeToMove);
        }

        public static T Deserialize<T>(string xml)
        {
            System.Xml.Serialization.XmlSerializer serializer =
                new System.Xml.Serialization.XmlSerializer(typeof(T));
            System.IO.StringReader stringReader = null;
            try
            {
                stringReader = new System.IO.StringReader(xml);
                return ((T)(serializer.Deserialize(System.Xml.XmlReader.Create(stringReader))));
            }
            finally
            {
                if ((stringReader != null))
                {
                    stringReader.Dispose();
                }
            }
        }

        public static T ReadXML<T>(string filePath)
        {
            T result = default(T);
            try
            {
                if (System.IO.File.Exists(filePath))
                {
                    XmlSerializer ser = new XmlSerializer(typeof(T));
                    using (FileStream myFileStream = new FileStream(filePath, FileMode.Open))
                    {
                        result = (T)ser.Deserialize(myFileStream);
                    }
                }
            }
            finally
            {
            }
            return result;
        }
        public static void WriteXML<T>(T data, string xmlFilePath)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                using (TextWriter tw = new StreamWriter(xmlFilePath))
                {
                    serializer.Serialize(tw, data);
                    tw.Close();
                }
            }
            finally
            {
            }
        }

        public static string Serialize(Object obj)
        {
            System.IO.StreamReader streamReader = null;
            System.IO.MemoryStream memoryStream = null;
            System.Xml.Serialization.XmlSerializer serializer =
                new System.Xml.Serialization.XmlSerializer(obj.GetType());
            try
            {
                memoryStream = new System.IO.MemoryStream();
                serializer.Serialize(memoryStream, obj);
                memoryStream.Seek(0, System.IO.SeekOrigin.Begin);
                streamReader = new System.IO.StreamReader(memoryStream);
                return streamReader.ReadToEnd();
            }
            finally
            {
                if ((streamReader != null))
                {
                    streamReader.Dispose();
                }
                if ((memoryStream != null))
                {
                    memoryStream.Dispose();
                }
            }
        }
    }
}
