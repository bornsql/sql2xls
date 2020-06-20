namespace SqlExcelExporter.Entities
{
    /// <summary>
    /// Container for configuration settings from the config.json file
    /// </summary>
    public class ConfigurationEntity
    {
        public string AdHocScriptFolder { get; set; }
        public string StoredProcedureFolder { get; set; }
        public string DefaultDatabase { get; set; }
        public bool UseDatabaseWhitelist { get; set; }
        public string[] DatabaseWhitelist { get; set; }
        public string ExportFilePathMac { get; set; }
        public string ExportFilePathWin { get; set; }
    }
}

