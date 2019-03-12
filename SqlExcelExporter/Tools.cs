using System.IO;
using Newtonsoft.Json;

namespace SqlExcelExporter
{
    public static class Tools
    {
        // 
        /// <summary>
        /// Deserialises a JSON file into an object of type T
        /// </summary>
        /// <typeparam name="T">The type of object</typeparam>
        /// <param name="file">The input file</param>
        /// <returns>The object as type T</returns>
        public static T ReadJsonItem<T>(FileInfo file)
        {
            if (!File.Exists(file.FullName))
            {
                return default(T);
            }

            var deserialiser = new JsonSerializer();

            var reader = new JsonTextReader(new StreamReader(file.FullName));
            return deserialiser.Deserialize<T>(reader);
        }

        /// <summary>
        /// Checks if assembly is running on Windows
        /// </summary>
        /// <returns>True or false</returns>
        public static bool IsWindows()
        {
            return Path.DirectorySeparatorChar.Equals('\\');
        }
    }
}