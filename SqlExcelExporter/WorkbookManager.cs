using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using SqlExcelExporter.Entities;

namespace SqlExcelExporter
{
    public class WorkbookManager
    {
        private readonly List<ResultEntity> m_instanceResults;
        private readonly Dictionary<string, List<ResultEntity>> m_databaseResults;

        public string FilePrefix { get; private set; }

        public WorkbookManager(List<ResultEntity> results)
        {
            FilePrefix = RemoveSpecialCharacters(Guid.NewGuid().ToString()).Trim();

            m_instanceResults = results.FindAll(x => x.Database.Equals("master", StringComparison.InvariantCultureIgnoreCase));

            m_databaseResults = new Dictionary<string, List<ResultEntity>>();

            var dbs = new List<string>();

            foreach (var d in results)
            {
                if (!dbs.Contains(d.Database) && !d.Database.Equals("master", StringComparison.InvariantCultureIgnoreCase))
                {
                    dbs.Add(d.Database);
                }
            }

            foreach (var database in dbs)
            {
                m_databaseResults.Add(database, results.FindAll(x => x.Database.Equals(database, StringComparison.InvariantCultureIgnoreCase)));
            }
        }

        private static string RemoveSpecialCharacters(string str)
        {
            var sb = new StringBuilder();
            foreach (var c in str.Where(c => char.IsLetterOrDigit(c) || c == '.' || c == '_'))
            {
                sb.Append(c);
            }
            return sb.ToString();
        }

        private static string GetMaxLengthSheetName(string workbook)
        {
            var w = RemoveSpecialCharacters(workbook).Trim();

            return w.Length > 31 ? w[..31] : w;
        }

        public FileInfo PrepareInstanceWorkbook()
        {
            var filename = new FileInfo($"{FilePrefix}_instance.xlsx");

            var ctr = 0;

            var workbook = new XLWorkbook();

            foreach (var dr in m_instanceResults)
            {
                Console.WriteLine($"Writing instance-level worksheet {ctr++}...");

                var worksheetName = GetMaxLengthSheetName(dr.Name);

                if (workbook.Worksheets.TryGetWorksheet(worksheetName, out _))
                {
                    var s = ctr.ToString().Trim();

                    worksheetName += s;
                    if (worksheetName.Length > 31)
                    {
                        worksheetName = worksheetName[..(31 - s.Length)] + s;
                    }
                }

                var worksheet = workbook.Worksheets.Add(worksheetName);

                if (dr.Results != null && dr.Results.Rows.Count > 0)
                {
                    for (var c = 0; c < dr.Results.Columns.Count; c++)
                    {
                        worksheet.Cell(1, c + 1).Value = dr.Results.Columns[c].ToString().Trim();
                    }

                    for (var r = 0; r < dr.Results.Rows.Count; r++)
                    {
                        for (var c = 0; c < dr.Results.Columns.Count; c++)
                        {
                            var cellValue = dr.Results.Rows[r][c].ToString()?.Trim();
                            if (cellValue != null && cellValue.Length > 32767)
                            {
                                cellValue = cellValue[..32767];
                            }

                            worksheet.Cell(r + 2, c + 1).Value = cellValue;
                        }
                    }

                    worksheet.Range(1, 1, dr.Results.Rows.Count + 1, dr.Results.Columns.Count).CreateTable();

                    var rngHeaders = worksheet.Range(1, 1, 1, dr.Results.Columns.Count); // The address is relative to rngTable (NOT the worksheet)
                    rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rngHeaders.Style.Font.Bold = true;
                    rngHeaders.Style.Fill.BackgroundColor = XLColor.Black;
                    rngHeaders.Style.Font.FontColor = XLColor.White;

                    worksheet.Columns(1, dr.Results.Columns.Count).AdjustToContents();

                    foreach (var column in worksheet.Columns(1, dr.Results.Columns.Count))
                    {
                        if (column.Width > 254)
                        {
                            column.Width = 254;
                        }
                    }
                }
                else
                {
                    worksheet.Cell("A1").Value = dr.Messages;
                }
            }

            Console.WriteLine("Saving output...");

            workbook.SaveAs(filename.FullName);

            return filename;
        }

        public FileInfo PrepareExcel(List<ResultEntity> results, ConfigurationEntity config)
        {
            var database = results.First().Database;

            // Skip things like ReportServer
            if (database.Equals("ReportServerTempDB", StringComparison.InvariantCultureIgnoreCase)
                || (config.UseDatabaseWhitelist
                    && !config.DatabaseWhitelist.Any(s => database.Equals(s, StringComparison.InvariantCultureIgnoreCase))))
            {
                return null;
            }

            var path = Tools.IsWindows() ? config.ExportFilePathWin : config.ExportFilePathMac;

            var filename = new FileInfo($"{path}{Path.DirectorySeparatorChar}{FilePrefix}_{RemoveSpecialCharacters(database).Trim()}.xlsx");
            var workbook = new XLWorkbook();

            var ctr = 0;

            foreach (var dr in results)
            {
                Console.WriteLine($"Writing worksheet {ctr++}...");

                var worksheetName = GetMaxLengthSheetName(dr.Name);

                if (workbook.Worksheets.TryGetWorksheet(worksheetName, out IXLWorksheet _))
                {
                    var s = ctr.ToString().Trim();

                    worksheetName += s;
                    if (worksheetName.Length > 31)
                    {
                        worksheetName = worksheetName[..(31 - s.Length)] + s;
                    }
                }

                var worksheet = workbook.Worksheets.Add(worksheetName);

                if (dr.Results != null && dr.Results.Rows.Count > 0)
                {
                    for (var c = 0; c < dr.Results.Columns.Count; c++)
                    {
                        worksheet.Cell(1, c + 1).Value = dr.Results.Columns[c].ToString().Trim();
                    }

                    for (var r = 0; r < dr.Results.Rows.Count; r++)
                    {
                        for (var c = 0; c < dr.Results.Columns.Count; c++)
                        {
                            var cellValue = dr.Results.Rows[r][c].ToString()?.Trim();
                            if (cellValue.Length > 32767)
                            {
                                cellValue = cellValue[..32767];
                            }

                            worksheet.Cell(r + 2, c + 1).Value = cellValue;
                        }
                    }

                    worksheet.Range(1, 1, dr.Results.Rows.Count + 1, dr.Results.Columns.Count).CreateTable();

                    var rngHeaders = worksheet.Range(1, 1, 1, dr.Results.Columns.Count); // The address is relative to rngTable (NOT the worksheet)
                    rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rngHeaders.Style.Font.Bold = true;
                    rngHeaders.Style.Fill.BackgroundColor = XLColor.Black;
                    rngHeaders.Style.Font.FontColor = XLColor.White;

                    worksheet.Columns(1, dr.Results.Columns.Count).AdjustToContents();

                    foreach (var column in worksheet.Columns(1, dr.Results.Columns.Count))
                    {
                        if (column.Width > 254)
                        {
                            column.Width = 254;
                        }
                    }
                }
                else
                {
                    worksheet.Cell("A1").Value = dr.Messages;
                }
            }

            Console.WriteLine("Saving output...");

            workbook.SaveAs(filename.FullName);

            return filename;
        }
    }
}
