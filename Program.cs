using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;

namespace Tab2Excel
{
	internal class Program
	{
		static ExcelPackage excelApp = null;
		static ExcelWorkbook excelWorkbook = null;
		static ExcelWorksheet worksheet = null;

		static void Main(string[] args)
		{
			if (args.Length == 0 
			|| (args.Length == 1 && args[0].Equals("help", StringComparison.OrdinalIgnoreCase))
			|| (args.Length == 1 && args[0].Equals("-help", StringComparison.OrdinalIgnoreCase)))
			{
				DisplayHelpMessage();
				return;
			}

			if (args.Length < 6) throw new Exception("Arguments must include \"{query}\" \"{outputFile}\" -s {server Name} -d {database Name}");

			Parameters parameters = ParseParameters(args);
			try
			{
				CreateExcelObjects();
				string connectionString = CreateConnectionString(parameters);

				using (SqlConnection connection = new SqlConnection(connectionString))
				{
					SqlCommand command = new SqlCommand(parameters.SourceQuery, connection);
					connection.Open();

					SqlDataReader reader = command.ExecuteReader();
					int startRow = 0;
					if (parameters.IncludeHeaderRow)
					{
						startRow = 1;
						AddHeaderRow(reader);
					}

					AddDataRows(startRow, reader);
					reader.Close();
				}

				FileInfo fileInfo = new FileInfo(parameters.OutputFilePath);
				excelApp.SaveAs(fileInfo);
			}
			finally
			{
				DisposeExcelObjects();
			}
		}

		static void DisplayHelpMessage()
		{
			Console.WriteLine("Description: Creates a Microsoft Excel file (xlsx) from a SQL statement result.");
			Console.WriteLine("----------------------------------------------------------------------------------------------------");
			Console.WriteLine("Usage:");
			Console.WriteLine("bcp2excel \"{sql query}\" \"{output file}\" -s {server name} -d {database name} [options]");
			Console.WriteLine("");
			Console.WriteLine("[options]:");
			Console.WriteLine("-u {user name} -p {password}");
			Console.WriteLine("Optional. Specifies a user name and password for the database connection. If not provided,the ");
			Console.WriteLine("application will attempt to use trusted connection credentials of the executing user account.");
			Console.WriteLine("");
			Console.WriteLine("-ch");
			Console.WriteLine("Optional. Specify to have the application include a header row containing the query column names.");
		}

		static Parameters ParseParameters(string[] args)
		{
			Parameters parameters = new Parameters();
			try
			{
				parameters.SourceQuery = args[0];
				parameters.OutputFilePath = args[1];
				parameters.OutputFilePath = Path.ChangeExtension(parameters.OutputFilePath, "xlsx");

				int index = args.ToList().IndexOf("-s");
				parameters.ServerName = args[index + 1];

				index = args.ToList().IndexOf("-d");
				parameters.DatabaseName = args[index + 1];

				index = args.ToList().IndexOf("-u");
				if (index == -1)
				{
					parameters.UseTrustedConnection = true;
				}
				else
				{
					parameters.UserName = args[index + 1];
					index = args.ToList().IndexOf("-p");
					parameters.Password = args[index + 1];
				}

				index = args.ToList().IndexOf("-ch");
				parameters.IncludeHeaderRow = (index > -1);
			}
			catch
			{
				throw new Exception("Specified parameters are invalid.  Check -help option for details.");
			}

			return parameters;
		}

		static void CreateExcelObjects()
		{
			excelApp = new ExcelPackage();
			excelWorkbook = excelApp.Workbook;
			worksheet = excelWorkbook.Worksheets.Add("Sheet1");
		}

		static void DisposeExcelObjects()
		{
			worksheet?.Dispose();
			excelWorkbook?.Dispose();
			excelApp?.Dispose();
			worksheet = null;
			excelWorkbook = null;
			excelApp = null;
		}

		static string CreateConnectionString(Parameters parameters)
		{
			if (parameters.UseTrustedConnection)
			{
				return $"Server={parameters.ServerName};Database={parameters.DatabaseName};Trusted_Connection=True;";
			}
			else
			{
				return $"Server={parameters.ServerName};Database={parameters.DatabaseName};User Id={parameters.UserName};Password={parameters.Password};";
			}
		}

		static void AddHeaderRow(SqlDataReader reader)
		{
			string[] columnNames =  reader
				.GetSchemaTable()
				.Rows
				.OfType<DataRow>()
				.Select(row => new { ColumnName = row.Field<string>("ColumnName"), ColumnOrdinal = row.Field <int>("ColumnOrdinal") })
				.OrderBy(c => c.ColumnOrdinal)
				.Select(c => c.ColumnName)
				.ToArray();

			for (int col = 0; col < columnNames.Length; col++) 
			{
				worksheet.Cells[1, col + 1].Value = columnNames[col];
			}
		}
	
		static void AddDataRows(int startRow, SqlDataReader reader)
		{
			int rowNum = 1;
			string[] columnNames = reader
				.GetSchemaTable()
				.Rows
				.OfType<DataRow>()
				.Select(row => new { ColumnName = row.Field<string>("ColumnName"), ColumnOrdinal = row.Field<int>("ColumnOrdinal") })
				.OrderBy(c => c.ColumnOrdinal)
				.Select(c => c.ColumnName)
				.ToArray();

			while (reader.Read())
			{
				for (int col = 0; col < columnNames.Length; col++)
				{
					object value = reader.GetValue(col);
					worksheet.Cells[startRow + rowNum, col + 1].Value = value;
				}
				rowNum++;
			}
		}
	}

	internal class Parameters
	{
		public string SourceQuery { get; set; } = null;

		public string OutputFilePath { get; set; } = null;

		public string ServerName { get; set; } = null;

		public string DatabaseName { get; set; } = null;

		public bool UseTrustedConnection { get; set; } = false;

		public string UserName { get; set; } = null;

		public string Password { get; set; } = null;

		public bool IncludeHeaderRow { get; set; } = false;
	}
}
