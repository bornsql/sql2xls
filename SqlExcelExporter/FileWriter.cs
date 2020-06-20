using System.Collections.Generic;
using System.IO;
using SqlExcelExporter.Entities;

namespace SqlExcelExporter
{
    public class FileWriter
    {
        private ConfigurationEntity m_config;

        public FileWriter(ConfigurationEntity config)
        {
            m_config = config;
        }

        public enum FileType
        {
            Excel,
            Json
        }

        private readonly List<ResultEntity> m_diagnosticResults;

        public List<FileInfo> ProduceExcel()
        {
            var excel = new List<FileInfo>();

            var helper = new WorkbookManager(m_diagnosticResults);

            var instance = helper.PrepareInstanceWorkbook();
            excel.Add(instance);

            var workbooks = helper.PrepareDatabaseWorkbooks(m_config);
            excel.AddRange(workbooks);

            return excel;
        }
    }
}
