using System;
using System.Collections.Generic;
using System.IO;
using SqlExcelExporter;
using SqlExcelExporter.Entities;

namespace SqlToExcelExporter
{
    class Program
    {
        private static ConnectionEntity m_connectionSettings;
        private static ConfigurationEntity m_config;

        private static void Main()
        {
            Console.WriteLine("--- SQL to Excel Exporter v0.2.1");
            Console.WriteLine("--- Copyright (c) Born SQL");
            Console.WriteLine("--- Written by Randolph West and other contributors. https://bornsql.ca/.");
            Console.WriteLine();

            var path = Directory.GetCurrentDirectory();
            var errorConnection = false;
            var errorDiagnostics = false;
            string error;
            var results = new List<ResultEntity>();

            try
            {
                m_connectionSettings = Tools.ReadJsonItem<ConnectionEntity>(new FileInfo("connection.json"));
            }
            catch (Exception ex)
            {
                error = ex.Message;
                errorConnection = true;
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] *** ERROR *** An error occurred validating the connection.json file [{error}].");
            }

            try
            {
                m_config = Tools.ReadJsonItem<ConfigurationEntity>(new FileInfo(Path.Combine(path, "config.json")));
            }
            catch (Exception ex)
            {
                error = ex.Message;
                errorDiagnostics = true;
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] *** ERROR *** An error occurred validating the config.json file [{error}].");
            }

            if (errorConnection || errorDiagnostics)
            {
                Console.WriteLine("Exiting.");
                return;
            }

            var database = m_config.DefaultDatabase;

            // Default value
            if (string.IsNullOrEmpty(database))
            {
                database = "master";
            }

            var dbh = new DatabaseHelper(m_connectionSettings);

            // Test connection
            if (!dbh.TestConnection())
            {
                return;
            }

            // Find stored procedure files
            var sp = new DirectoryInfo(m_config.StoredProcedureFolder);

            if (!string.IsNullOrEmpty(sp.Name) && Directory.Exists(sp.FullName))
            {
                foreach (var file in sp.GetFiles("*.sql", SearchOption.TopDirectoryOnly))
                {
                    Console.WriteLine($"Installing stored procedure from file [{file.Name}] in database [{database}]...");
                    dbh.InstallStoredProcedure(database, file);
                }
            }

            // Find ad hoc scripts
            var di = new DirectoryInfo(m_config.AdHocScriptFolder);

            if (string.IsNullOrEmpty(di.Name) || !Directory.Exists(di.FullName))
            {
                Console.WriteLine("No script files, or path does not exist.");
            }
            else
            {
                foreach (var fi in di.GetFiles("*.sql", SearchOption.TopDirectoryOnly))
                {
                    Console.WriteLine($"Running script [{fi.Name}] in database [{database}]");
                    results.AddRange(dbh.RunStandaloneScript(database, fi.Name.Replace(fi.Extension, ""), fi));
                }
            }

            var mgr = new WorkbookManager(results);
            var f = mgr.PrepareExcel(results, m_config);

            Console.WriteLine(
                $"Excel file [{f.Name}] generated in folder [{(Tools.IsWindows() ? m_config.ExportFilePathWin : m_config.ExportFilePathMac)}].");

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
