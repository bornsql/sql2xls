using System.Collections.Generic;
using System.IO;

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

        private const string PathName = @"C:\Temp";

        private List<ResultEntity> m_diagnosticResults;

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
