namespace SqlExcelExporter.Entities
{
    /// <summary>
    /// Container for connection settings from the connection.json file
    /// </summary>
    public class ConnectionEntity
    {
        /// <summary>
        /// Generates a data source string based on instance and port settings.
        /// </summary>
        private string DataSource => !string.IsNullOrEmpty(InstanceName)
            ? $"{ServerName}\\{InstanceName}"
            : (Port != 1433 ? $"{ServerName},{Port}" : $"{ServerName}");

        /// <summary>
        /// Returns true if the instance is a named instance.
        /// </summary>
        public bool IsNamedInstance => !string.IsNullOrEmpty(InstanceName);

        /// <summary>
        /// The SQL Server host name.
        /// </summary>
        public string ServerName { get; set; }


        /// <summary>
        /// The SQL Server instance name. A default instance is usually left blank.
        /// </summary>
        public string InstanceName { get; set; }

        /// <summary>
        /// The SQL Server instance port. Defaults to 1433.
        /// </summary>
        public int Port { get; set; }

        /// <summary>
        /// The default database to connect to. Defaults to "master".
        /// </summary>
        private const string InitialCatalog = "master";

        /// <summary>
        /// Returns true if using Integrated Authentication. If using SQL Server Authentication, this will be false.
        /// </summary>
        public bool UseTrustedConnection { get; set; }

        /// <summary>
        /// The user name if using SQL Server Authentication.
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// The password (stored in clear text) if using SQL Server Authentication. It is not secure because it is passed around in the connection string.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// The connection timeout in seconds. Defaults to 0 for unlimited timeout.
        /// </summary>
        private const int ConnectionTimeout = 0;

        /// <summary>
        /// The application name that appears when querying active connections on the SQL Server instance.
        /// </summary>
        private const string ApplicationName = "SqlToExcelExporter by bornsql.ca";

        /// <summary>
        /// A complete connection string for passing to the application
        /// </summary>
        /// <returns>The connection string including credentials, initial catalogue and data source</returns>
        public override string ToString()
        {
            return UseTrustedConnection
                ? $@"Data Source={DataSource};Initial Catalog={InitialCatalog};Trusted_Connection=True;Connection Timeout={ConnectionTimeout};Application Name={ApplicationName}"
                : $@"Data Source={DataSource};Initial Catalog={InitialCatalog};User Id={UserId};Password={Password};Connection Timeout={ConnectionTimeout};Application Name={ApplicationName}";
        }
    }
}