using System;
using System.Collections.Generic;
using System.Text;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Collections;

namespace Com.Bing
{
    public static class UnNetZHelper
    {
        public static MemoryStream UnZip(byte[] data)
        {
            InflaterInputStream zipStream = new InflaterInputStream(new MemoryStream(data));
            byte[] buffer = new byte[data.Length];
            MemoryStream upZipStream = new MemoryStream();
            while (true)
            {
                int num = zipStream.Read(buffer, 0, buffer.Length);
                if (num <= 0) break;
                upZipStream.Write(buffer, 0, num);
            }
            upZipStream.Flush();
            upZipStream.Seek(0, SeekOrigin.Begin);
            return upZipStream;
        }
        public static string UnMangleDllName(string str)
        {
            return str.Replace("!1", " ").Replace("!2", ",").Replace("!3", ".Resources").Replace("!3", ".resources").Replace("!4", "Culture");
        }
        public static Dictionary<string, byte[]> GetResourceDict(string filePath)
        {
            Dictionary<string, byte[]> dict = new Dictionary<string, byte[]>();

            byte[] bytes = File.ReadAllBytes(filePath);
            Assembly assembly = Assembly.Load(bytes);
            string[] resourcesName = assembly.GetManifestResourceNames();

            foreach (string resourceName in resourcesName)
            {
                using (System.IO.UnmanagedMemoryStream stream = (System.IO.UnmanagedMemoryStream)(assembly.GetManifestResourceStream(resourceName)))
                {
                    ResourceManager rm = new ResourceManager("app", assembly);
                    ResourceSet resourceSet = null;
                    try
                    {
                        resourceSet = rm.GetResourceSet(System.Threading.Thread.CurrentThread.CurrentCulture, true, true);
                    }
                    catch (MissingManifestResourceException)
                    {
                        return null;
                    }
                    foreach (DictionaryEntry entry in resourceSet)
                    {
                        string key = entry.Key.ToString();
                        byte[] contentBytes = entry.Value as byte[];
                        dict[key] = contentBytes;
                    }
                }
            }
            return dict;
        }
    }
}
