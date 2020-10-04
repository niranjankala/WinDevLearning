using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Simergy.Common
{
    public static class ExceptionExtension
    {
        public static IEnumerable<IEnumerable<T>> Split<T>(this IEnumerable<T> source, int parts)
        {
            var list = new List<T>(source);
            int defaultSize = (int)((double)list.Count / (double)parts);
            int offset = list.Count % parts;
            int position = 0;

            for (int i = 0; i < parts; i++)
            {
                int size = defaultSize;
                if (i < offset)
                    size++; // Just add one to the size (it's enough).

                yield return list.GetRange(position, size);

                // Set the new position after creating a part list, so that it always start with position zero on the first yield return above.
                position += size;
            }
        }


        public static void LogExceptionMessage(this Exception ex)
        {
            if (ex != null)
                ex.LogException();
        }

        public static void LogException(this Exception ex)
        {
            if (ex != null)
            {
                //Logging.LogManager.LogException(ex);
            }
        }
    }

    public static class ExtensionHelpers
    {
        public static bool IsNullOrDefault<T>(T value)
        {
            return object.Equals(value, default(T));
        }
    }

    public static class IDictionaryExtensions
    {
        public static TKey FindKeyByValue<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TValue value)
        {
            if (dictionary == null)
                throw new ArgumentNullException("dictionary");

            foreach (KeyValuePair<TKey, TValue> pair in dictionary)
                if (value.Equals(pair.Value)) return pair.Key;

            throw new Exception("the value is not found in the dictionary");
        }

        public static void AddRange<T, S>(this Dictionary<T, S> source, Dictionary<T, S> collection)
        {
            if (collection == null)
            {
                throw new ArgumentNullException("Collection is null");
            }

            foreach (var item in collection)
            {
                if (!source.ContainsKey(item.Key))
                {
                    source.Add(item.Key, item.Value);
                }
                else
                {
                    // handle duplicate key issue here
                }
            }
        }
    }

    public static class DateTimeExtensions
    {

        public static int GetDaysInYear(this DateTime dateTime)
        {
            int numberOfDaysInYear = -1;
            if (dateTime.Year >= DateTime.MinValue.Year && dateTime.Year <= DateTime.MaxValue.Year - 1)
            {
                var thisYear = new DateTime(dateTime.Year, 1, 1);
                var nextYear = new DateTime(dateTime.Year + 1, 1, 1);
                numberOfDaysInYear = (nextYear - thisYear).Days;
            }
            return numberOfDaysInYear;
        }
    }

    public static class RandomExtension
    {
        public static double NextDouble(this Random rnd, double min, double max)
        {
            return rnd.NextDouble() * (max - min) + min;
        }
    }

    public static class StringExtension
    {
        public static string UppercaseFirst(this string convertFrom)
        {
            // Check for empty string.
            if (string.IsNullOrEmpty(convertFrom))
            {
                return string.Empty;
            }
            // Return char and concat substring.
            return char.ToUpper(convertFrom[0]) + convertFrom.Substring(1);
        }

        public static string ReplaceLastOccurrence(string Source, string Find, string Replace)
        {
            int place = Source.LastIndexOf(Find);

            if (place == -1)
                return Source;

            string result = Source.Remove(place, Find.Length).Insert(place, Replace);
            return result;
        }
    }

    public static class SafeConversion
    {
        public static double? ToDouble(string inVal)
        {
            double retVal;
            if (!double.TryParse(inVal, out retVal))
            {
                return null;
            }
            return retVal;
        }
    }

    public static class TPLExtensions
    {
        public static void RaiseCancellationIfRequested(System.Threading.CancellationToken ct)
        {
            if (ct != null && ct.IsCancellationRequested)
                ct.ThrowIfCancellationRequested();
        }
    }
    public static class XMLExtensions
    {
        public static T ConvertNode<T>(this XmlNode node, string rootAttribute) where T : class
        {
            MemoryStream stm = new MemoryStream();

            StreamWriter stw = new StreamWriter(stm);
            stw.Write(node.OuterXml);
            stw.Flush();

            stm.Position = 0;

            XmlSerializer ser = new XmlSerializer(typeof(T));
            if (!string.IsNullOrWhiteSpace(rootAttribute))
                ser = new XmlSerializer(typeof(T), new XmlRootAttribute("rootAttribute"));
            T result = (ser.Deserialize(stm) as T);

            return result;
        }

        public static T ConvertNode<T>(this XmlNode node) where T : class
        {
            MemoryStream stm = new MemoryStream();

            StreamWriter stw = new StreamWriter(stm);
            stw.Write(node.OuterXml);
            stw.Flush();

            stm.Position = 0;

            XmlSerializer ser = new XmlSerializer(typeof(T));
            T result = (ser.Deserialize(stm) as T);

            return result;
        }
    }

    public static class Comparison
    {
        public static bool PublicInstancePropertiesEqual<T>(T self, T to, params string[] ignore) where T : class
        {
            if (self != null && to != null)
            {
                Type type = typeof(T);
                List<string> ignoreList = new List<string>(ignore);
                foreach (System.Reflection.PropertyInfo pi in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
                {
                    if (!ignoreList.Contains(pi.Name))
                    {
                        object selfValue = type.GetProperty(pi.Name).GetValue(self, null);
                        object toValue = type.GetProperty(pi.Name).GetValue(to, null);

                        if (selfValue != toValue && (selfValue == null || !selfValue.Equals(toValue)))
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            return self == to;
        }
    }

    public static class ObjectExtension
    {
        public static T DeepClone<T>(T original)
        {
            if (!typeof(T).IsSerializable)
            {
                throw new ArgumentException("The type must be serializable.", "original");
            }

            if (ReferenceEquals(original, null))
            {
                return default(T);
            }

            using (var stream = new MemoryStream())
            {
                var formatter = new BinaryFormatter
                {
                    Context = new StreamingContext(StreamingContextStates.Clone)
                };

                formatter.Serialize(stream, original);
                stream.Position = 0;

                return (T)formatter.Deserialize(stream);
            }
        }
    }

    public static class ListExtension
    {
        public static string ConvertToCSVString(this IList _valuesList, string separator)
        {
            string finalvalue = string.Empty;
            if (_valuesList != null && _valuesList.Count > 0)
            {
                for (int cnt = 0; cnt <= _valuesList.Count - 1; cnt++)
                {
                    if (cnt > 0)
                    {
                        finalvalue += separator + Convert.ToString(_valuesList[cnt]);
                    }
                    else
                    {
                        finalvalue = Convert.ToString(_valuesList[cnt]);
                    }

                }
            }
            return finalvalue;
        }
    }
}
