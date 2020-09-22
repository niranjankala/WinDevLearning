using System.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace WinDev.Common.Helpers
{
    public class JSONHelper
    {
        public static T DownloadAndDeserializeJsonData<T>(string url) where T : new()
        {
            try
            {
                using (var webClient = new WebClient())
                {
                    var jsonData = string.Empty;
                    try
                    {
                        jsonData = webClient.DownloadString(url);
                    }
                    catch (Exception) { }

                    return DeserializeJsonData<T>(jsonData);
                }
            }
            catch
            {
                return DeserializeJsonData<T>(null); 
            }
        }

        public static T DeserializeJsonData<T>(string jsonData) where T : new()
        {
            return !string.IsNullOrEmpty(jsonData)
                           ? JsonConvert.DeserializeObject<T>(jsonData)
                           : new T();
        }
    }
}
