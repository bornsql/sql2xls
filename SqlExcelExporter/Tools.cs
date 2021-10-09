using System.IO;
using System.Text.Json;

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
            return !File.Exists(file.FullName)
                ? default
                : JsonSerializer.Deserialize<T>(File.ReadAllText(file.FullName));
        }

        /// <summary>
        /// Checks if assembly is running on Windows
        /// </summary>
        /// <returns>True or false</returns>
        public static bool IsWindows() => Path.DirectorySeparatorChar.Equals('\\');
    }
}