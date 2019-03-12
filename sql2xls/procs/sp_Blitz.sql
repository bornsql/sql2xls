IF OBJECT_ID('dbo.sp_Blitz') IS NULL
  EXEC ('CREATE PROCEDURE dbo.sp_Blitz AS RETURN 0;');
GO
/*
--Sample execution call with the most common parameters:
EXEC [dbo].[sp_Blitz]
    @CheckUserDatabaseObjects = 1 ,
    @CheckProcedureCache = 0 ,
    @OutputType = 'TABLE' ,
    @OutputProcedureCache = 0 ,
    @CheckProcedureCacheFilter = NULL,
    @CheckServerInfo = 1
*/
