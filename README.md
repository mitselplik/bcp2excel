# bcp2excel
A command-line tool for creating a Microsoft Excel file (xlsx) from a SQL statement result.

## Usage:
bcp2excel "{sql query}" "{output file}" -s {server name} -d {database name} [options]

### [options]
-u {user name} -p {password}
Optional. Specifies a user name and password for the database connection. If not provided,the application will attempt to use trusted connection credentials of the executing user account.

-ch
Optional. Specify to have the application include a header row containing the query column names.

## Example:
```
bcp2excel "select * from orders" "\\my_network_share\account.xlsx" -s localhost -d contoso -ch
```

## Notes:
This may be used to call from a stored procedure using xp_cmdshell as well.  
You will need add the installation folder to your *path* environment variable.
You will also need to enable use of xp_cmdshell from SQL Server.
Afterward, the syntax to call the program 

### Example:

```SQL
DECLARE @sql_string nvarchar(max) = N'select * from orders' -- Note: Query *must* be all on one line.
DECLARE @export_path nvarchar(max) = N'\\my_network_share\account.xlsx'
DECLARE @cmd_sql nvarchar(max) = N'EXEC xp_cmdshell ''bcp2excel "' + @sql_string + '" "' + @export_path + '" -s ' + @@SERVERNAME + ' -d ' + DB_NAME() + ' -ch'', no_output'
EXEC sp_executesql @cmd_sql
```



