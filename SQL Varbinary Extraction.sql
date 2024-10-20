
/*
sp_configure 'show advanced options', 1;

GO

RECONFIGURE;

GO

sp_configure 'Ole Automation Procedures', 1;

GO

RECONFIGURE;

GO

sp_configure
 */




sp_configure 'Ole Automation Procedures', 1

 
DECLARE 
@OUTPUT_PATH VARCHAR(MAX), 
@FILE_PATH VARCHAR(MAX),
@FILE_DATA VARBINARY(MAX),
@DOCUMENT_ID VARCHAR(MAX),
@FILEEXTENSION VARCHAR(MAX),
@ObjectToken INT
DECLARE EXPORT_DATA CURSOR FAST_FORWARD FOR
 
--**Start Query Section of Script**--
 
SELECT EventID
	  ,cast(replace(cast(Text as varchar(max)),left(cast(text as varchar(max)),charindex('</head>',cast(text as varchar(max)),1)+6),'<html>') as varbinary(max)) as data
	  ,concat('',,'\') as filepath
	  ,'.html'
  FROM [Event]
  where  FinalText is null and Text not like '%base64%'
 
--**End Query Section of Script**--
 
OPEN EXPORT_DATA
FETCH NEXT FROM EXPORT_DATA INTO @DOCUMENT_ID, @FILE_DATA, @OUTPUT_PATH, @FILEEXTENSION
WHILE @@FETCH_STATUS = 0
BEGIN
EXEC xp_create_subdir @OUTPUT_PATH;
SET @FILE_PATH = @OUTPUT_PATH + @DOCUMENT_ID + @FILEEXTENSION
PRINT @FILE_PATH
EXEC sp_OACreate 'ADODB.Stream', @ObjectToken OUTPUT
EXEC sp_OASetProperty @ObjectToken, 'Type', 1
EXEC sp_OAMethod @ObjectToken, 'Open'
EXEC sp_OAMethod @ObjectToken, 'Write', NULL, @FILE_DATA
EXEC sp_OAMethod @ObjectToken, 'SaveToFile', NULL, @FILE_PATH, 2
EXEC sp_OAMethod @ObjectToken, 'Close'
EXEC sp_OADestroy @ObjectToken
FETCH NEXT FROM EXPORT_DATA INTO @DOCUMENT_ID, @FILE_DATA, @OUTPUT_PATH, @FILEEXTENSION
END
CLOSE EXPORT_DATA
DEALLOCATE EXPORT_DATA
PRINT 'Export complete.'