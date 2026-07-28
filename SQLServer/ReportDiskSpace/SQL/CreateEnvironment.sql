USE [DBAdmin]
GO

/****** Object:  Table [REP].[DestinationReport]    Script Date: 5/18/2026 5:01:41 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [REP].[DestinationReport](
	[DatabaseName] [nvarchar](128) NOT NULL,
	[InstanceName] [nvarchar](100) NOT NULL,
	[GroupName] [nvarchar](100) NOT NULL,
	[DestMail] [nvarchar](100) NOT NULL
) ON [DBAdmin_Cold_FG_01]
GO


USE [DBAdmin]
GO

/****** Object:  View [REP].[TopDatabasesSpaceUsage]    Script Date: 5/18/2026 5:01:58 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




CREATE VIEW [REP].[TopDatabasesSpaceUsage]
AS


SELECT 
	  [DatabaseName]
    , [TotalSize]
    , [AvailableSize]
    , [QueryDate]
    , [InstanceName]
FROM
(
	SELECT  
		  [DatabaseName]
		, [TotalSize]
		, [AvailableSize]
		, [QueryDate]
		, [InstanceName]
		, ROW_NUMBER() OVER (PARTITION BY DatabaseName, InstanceName, SUBSTRING(CAST(QueryDate AS VARCHAR),1,8) ORDER BY QueryDate ASC) AS RowNum
	FROM
	(
		SELECT  
			  [DatabaseName]
			, MAX([TotalSize]) AS [TotalSize]
			, MAX([AvailableSize]) AS [AvailableSize] 
			, [QueryDate]
			, [InstanceName]		 
		FROM 
		(
			SELECT 
				  [Database Name] AS [DatabaseName]     
				, SUM([Total Size in MB]) AS TotalSize
				, SUM([Available Space In MB]) AS AvailableSize      
				, CAST([QueryDateTime] AS DATE) AS QueryDate	  
				, [ServerName]
				, [InstanceName]
			FROM [DBAdmin].[DataWarehouse].[FileSizes]  
			WHERE [Filegroup Name] <> ''
			AND InstanceName like '%PROD%'
			AND [Database Name] IN (
				'ADV'
			, 'RiskShield'
			, 'IssuingCurated'
			, 'AurinDM_DP'
			, 'IssuingData'
			) 
			GROUP BY [Database Name], [ServerName], [InstanceName], CAST([QueryDateTime] AS DATE)
		) tblA
		GROUP BY DatabaseName , QueryDate, InstanceName
	) tblB
) tblC
WHERE tblC.RowNum = 1
AND QueryDate > DATEADD(m,-6,GETDATE())
GO




USE [DBAdmin]
GO

/****** Object:  View [REP].[ReportDiskTopDatabases]    Script Date: 5/18/2026 5:01:48 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE VIEW [REP].[ReportDiskTopDatabases] AS
WITH Months AS (
    SELECT CONVERT(varchar(10), CONCAT(YEAR(DATEADD(m,-5,GETDATE())),'-',RIGHT('0' + CAST(MONTH(DATEADD(m,-5,GETDATE())) AS VARCHAR(2)),2)), 23) AS MonthVal UNION ALL 
	SELECT CONVERT(varchar(10), CONCAT(YEAR(DATEADD(m,-4,GETDATE())),'-',RIGHT('0' + CAST(MONTH(DATEADD(m,-4,GETDATE())) AS VARCHAR(2)),2)), 23) UNION ALL 
	SELECT CONVERT(varchar(10), CONCAT(YEAR(DATEADD(m,-3,GETDATE())),'-',RIGHT('0' + CAST(MONTH(DATEADD(m,-3,GETDATE())) AS VARCHAR(2)),2)), 23) UNION ALL
    SELECT CONVERT(varchar(10), CONCAT(YEAR(DATEADD(m,-2,GETDATE())),'-',RIGHT('0' + CAST(MONTH(DATEADD(m,-2,GETDATE())) AS VARCHAR(2)),2)), 23) UNION ALL 
	SELECT CONVERT(varchar(10), CONCAT(YEAR(DATEADD(m,-1,GETDATE())),'-',RIGHT('0' + CAST(MONTH(DATEADD(m,-1,GETDATE())) AS VARCHAR(2)),2)), 23) UNION ALL 
	SELECT CONVERT(varchar(10), CONCAT(YEAR(DATEADD(m,-0,GETDATE())),'-',RIGHT('0' + CAST(MONTH(DATEADD(m,-0,GETDATE())) AS VARCHAR(2)),2)), 23)
),
DbInstanceList AS (
	SELECT DISTINCT
		  DatabaseName
		, InstanceName
	 FROM [DBAdmin].[REP].[TopDatabasesSpaceUsage]
),
DateAndDbs AS (
	SELECT
		  MonthVal
		, DatabaseName AS DBName
		, InstanceName AS InstName
	FROM Months 
	CROSS APPLY DbInstanceList
)
SELECT 
	  ISNULL(DatabaseName,DBName) AS DatabaseName
	, ISNULL(InstanceName,InstName) AS InstanceName
	, ISNULL(AvailableSize,0) AS AvailableSize
	, ISNULL(TotalSize,0) AS TotalSize
	, ISNULL(QueryDate,CAST(CONCAT(MonthVal,'-01') AS DATE)) AS QueryDate
FROM DateAndDbs
LEFT OUTER JOIN [DBAdmin].[REP].[TopDatabasesSpaceUsage] us
ON DateAndDbs.MonthVal = CONVERT(varchar(10), CONCAT(YEAR(us.QueryDate),'-',RIGHT('0' + CAST(MONTH(us.QueryDate) AS VARCHAR(2)),2)), 23)
AND DateAndDbs.DBName = DatabaseName 
AND DateAndDbs.InstName = InstanceName
GO

