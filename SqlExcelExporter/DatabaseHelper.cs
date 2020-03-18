using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using System.Text;

namespace SqlExcelExporter
{
    public class DatabaseHelper
    {
        private readonly ConnectionEntity m_connEntity;

        public DatabaseHelper(ConnectionEntity connEntity)
        {
            m_connEntity = connEntity;
        }

        /// <summary>
        /// Tests that the target SQL Server can be reached.
        /// This method does not test the connection string, just the TCP endpoint.
        /// </summary>
        /// <returns>True or False</returns>
        public bool TestConnection()
        {
            try
            {
                var tcp = new TcpClient(m_connEntity.ServerName, m_connEntity.Port);
                Console.WriteLine(tcp.Connected ? $"[{DateTime.Now:HH:mm:ss}] *** [INFO] Opened connection to {0}" : $"[{DateTime.Now:HH:mm:ss}] *** [ERROR] {0} not connected", m_connEntity.ServerName);
                tcp.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] *** [ERROR] Cannot connect to [{m_connEntity.ServerName}] - please check connection settings: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Runs a standalone script against the master database
        /// This script can contain multiple batch separators (GO keyword)
        /// </summary>
        /// <param name="database">The target database name</param>
        /// <param name="scriptName">The name of the script for the Excel worksheet</param>
        /// <param name="script">The script file</param>
        /// <returns>A list of one or more results that will be parsed by the file writer</returns>
        public List<ResultEntity> RunStandaloneScript(string database, string scriptName, FileInfo script)
        {
            return File.Exists(script.FullName)
                ? RunStandaloneScript(m_connEntity.ToString(), database, scriptName, File.ReadAllText(script.FullName))
                : null;
        }

        private List<ResultEntity> RunStandaloneScript(string connectionString, string database, string scriptName, string script)
        {
            var d = new List<ResultEntity>();

            using (var cn = new SqlConnection(connectionString))
            {
                try
                {
                    if (!database.Equals("master", StringComparison.InvariantCultureIgnoreCase))
                    {
                        cn.ChangeDatabase(database);
                    }

                    var sqlBatch = string.Empty;
                    var cmd = new SqlCommand(string.Empty, cn);
                    cn.Open();

                    script += $"{Environment.NewLine}GO"; // make sure last batch is executed.

                    foreach (var line in script.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (line.Trim().Equals("GO", StringComparison.InvariantCultureIgnoreCase))
                        {
                            cmd.CommandText = sqlBatch;
                            var ds = new DataSet();

                            if (cmd.CommandText.Length > 0)
                            {
                                var da = new SqlDataAdapter
                                {
                                    SelectCommand = new SqlCommand(cmd.CommandText, cn)
                                    {
                                        CommandTimeout = 0
                                    }
                                };
                                da.Fill(ds);
                            }
                            sqlBatch = string.Empty;

                            var ctr = 0;

                            foreach (DataTable dt in ds.Tables)
                            {
                                d.Add(new ResultEntity
                                {
                                    Database = database,
                                    Duration = new TimeSpan(0),
                                    ErrorMessage = string.Empty,
                                    Results = dt,
                                    Messages = string.Empty,
                                    Name = $"{scriptName}{++ctr}",
                                    IsError = false
                                });
                            }
                        }
                        else
                        {
                            sqlBatch += line + $"{Environment.NewLine}";
                        }
                    }
                }
                finally
                {
                    cn.Close();
                }
            }
            return d;
        }

        /// <summary>
        /// Install a stored procedure to a target database from a script file
        /// </summary>
        /// <param name="database">The target database</param>
        /// <param name="script">The script file</param>
        public void InstallStoredProcedure(string database, FileInfo script)
        {
            if (File.Exists(script.FullName))
            {
                InstallStoredProcedure(m_connEntity.ToString(), database, File.ReadAllText(script.FullName));
            }
        }

        private void InstallStoredProcedure(string connectionString, string database, string script)
        {
            // Based on this answer: http://stackoverflow.com/a/40827
            using (var cn = new SqlConnection(connectionString))
            {
                try
                {
                    if (!database.Equals("master", StringComparison.InvariantCultureIgnoreCase))
                    {
                        cn.ChangeDatabase(database);
                    }

                    var sqlBatch = string.Empty;
                    var cmd = new SqlCommand(string.Empty, cn);
                    cn.Open();

                    script += $"{Environment.NewLine}GO"; // make sure last batch is executed.

                    foreach (var line in script.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (line.Trim().Equals("GO", StringComparison.InvariantCultureIgnoreCase))
                        {
                            cmd.CommandText = sqlBatch;

                            if (cmd.CommandText.Length > 0)
                            {
                                cmd.ExecuteNonQuery();
                            }
                            sqlBatch = string.Empty;
                        }
                        else
                        {
                            sqlBatch += line + $"{Environment.NewLine}";
                        }
                    }
                }
                finally
                {
                    cn.Close();
                }
            }
        }

        /// <summary>
        /// Runs a single ad hoc script against a target database
        /// </summary>
        /// <param name="database">The target database</param>
        /// <param name="scriptName">The name of the script for the Excel worksheet</param>
        /// <param name="script">The script file</param>
        /// <returns>A result that will be parsed by the file writer</returns>
        public ResultEntity RunScript(string database, string scriptName, FileInfo script)
        {
            return File.Exists(script.FullName)
                ? RunScript(m_connEntity.ToString(), database, scriptName, File.ReadAllText(script.FullName))
                : null;
        }

        private ResultEntity RunScript(string connectionString, string database, string scriptName, string script)
        {
            var results = new ResultEntity
            {
                Database = database,
                Name = scriptName,
                IsError = false
            };

            var dt = new DataTable();
            var da = new SqlDataAdapter();
            var sb = new StringBuilder();

            using (var cn = new SqlConnection(connectionString))
            {
                try
                {
                    cn.Open();

                    var timer = Stopwatch.StartNew();

                    if (!database.Equals("master", StringComparison.InvariantCultureIgnoreCase))
                    {
                        cn.ChangeDatabase(database);
                    }

                    cn.InfoMessage += delegate (object sender, SqlInfoMessageEventArgs e)
                    {
                        sb.AppendLine(e.Message);
                    };

                    var cmd = new SqlCommand(script, cn)
                    {
                        CommandType = CommandType.Text,
                        CommandTimeout = 0
                    };

                    da.SelectCommand = cmd;
                    da.Fill(dt);

                    dt.TableName = scriptName.Replace(" ", "");

                    results.Results = dt;
                    results.Messages = sb.ToString();
                    results.Duration = timer.Elapsed;

                    timer.Stop();

                    return results;
                }
                catch (Exception ex)
                {
                    results.IsError = true;
                    results.ErrorMessage = ex.Message;
                }
            }
            return results;
        }
    }
}
