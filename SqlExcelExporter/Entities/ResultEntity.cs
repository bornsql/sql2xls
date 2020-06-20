using System;
using System.Data;

namespace SqlExcelExporter.Entities
{
    /// <summary>
    /// Container for results from one or more queries against SQL Server
    /// </summary>
    public class ResultEntity
    {
        public string Database { get; set; }
        public string Name { get; set; }
        public DataTable Results { get; set; }
        public string Messages { get; set; }
        public bool IsError { get; set; }
        public string ErrorMessage { get; set; }
        public TimeSpan Duration { get; set; }
    }
}