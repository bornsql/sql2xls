# sql2xls - SQL to Excel Export Tool

Imagine being able to export the results of a T-SQL query straight into an Excel file. Imagine being able to export the results of an unlimited number of scripts at one time.

Well now you can with my free tool *sql2xls*. The latest beta v0.2.0 is available for download from this link. You can export the results of one or more scripts into an Excel file per instance or database.

The latest version runs on .NET Core 3.1. It has been tested on Windows 10 and macOS Catalina.

## Requirements

- .NET Core 3.1 SDK to compile and run the executable.
- `mono-libgdiplus` for macOS, or `libgdiplus` for Linux.

## Configuration

- Edit the connection.json file to set up a connection to a SQL Server instance.
- Edit the config.json file to change the path of the output folder.

## Stored procedures

Add any stored procedures you want to install, to the procs folder. Files must have a .sql extension.

## Ad hoc scripts

Add any scripts you want to execute, to the scripts folder. Files must have a .sql extension.

_Note: If you want to execute any of the stored procedures, you must add a file in the scripts folder which runs the stored procedure. Each stored procedure you need to run must be separated by the GO batch separator on its own line. It’s probably better to keep these in their own separate files though._

## Excel output

One file is created per database, including the master database.

## Dependencies

- ClosedXML (for generating the Excel output)
- Microsoft.Data.SqlClient (for connecting to SQL Server)

## Licence

- This project uses the MIT licence.
- ClosedXML is copyright (c) 2017 ClosedXML

## Future plans

- JSON export
- SQL Server Management Studio plugin

_THIS IS BETA SOFTWARE AND SHOULD NOT BE RUN ON A PRODUCTION ENVIRONMENT WITHOUT EXTENSIVE TESTING._
