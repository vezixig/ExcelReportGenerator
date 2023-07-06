namespace ReportGenerator;

using System.Data;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

/// <summary>Represents a report generator.</summary>
internal class ReportGenerator
{
    private const string AccountingFormat = "_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-";

    private const string DateFormat = "dd.MM.yyyy";

    private const string ConnectionName = "SqlConnection";

    private static IConfigurationRoot? _configuration;

    #region Methods

    /// <summary>Generates reports based on the provided report definitions.</summary>
    /// <param name="reports">The list of report definitions.</param>
    public void GenerateReport(List<ReportDefinition> reports)
    {
        foreach (var report in reports)
        {
            Console.WriteLine($"Starting generation of '{report.FilePath}'");

            using var workbook = new XLWorkbook();
            foreach (var page in report.Pages)
            {
                Console.WriteLine($"\tAdding page {page.Name}");

                Console.WriteLine("\t\tFetching data...");
                var data = GetData(page.Sql);
                Console.WriteLine("\t\tAdding worksheet...");
                AddSheet(workbook, data, page);
            }

            Console.WriteLine("\tSaving to file...");
            workbook.SaveAs(report.FilePath);
            Console.WriteLine("\tDone");
        }
    }

    private static void AddSheet(IXLWorkbook workbook, DataTable dataTable, ReportPage page)
    {
        var worksheet = workbook.Worksheets.Add(page.Name);

        // Add column headers
        for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
            worksheet.Cell(1, columnIndex + 1).Value = dataTable.Columns[columnIndex].ColumnName;

        // Add data rows
        for (var rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
        for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
        {
            var currentCell = worksheet.Cell(rowIndex + 2, columnIndex + 1);
            var dataValue = dataTable.Rows[rowIndex][columnIndex];
            page.SpecialColumns.TryGetValue(dataTable.Columns[columnIndex].ColumnName, out var columnType);

            SetCellValue(columnType, currentCell, dataValue);
            SetCellType(columnType, currentCell);
        }

        // Format as table
        var firstCell = worksheet.FirstCellUsed();
        var lastCell = worksheet.LastCellUsed();
        var range = worksheet.Range(firstCell.Address, lastCell.Address);
        var table = range.CreateTable();
        table.Theme = XLTableTheme.TableStyleMedium21;

        worksheet.Columns().AdjustToContents();
    }

    private static void SetCellValue(ColumnType columnType, IXLCell currentCell, object dataValue)
    {
        switch (columnType)
        {
            case ColumnType.Accounting:
                var numberFormat = new NumberFormatInfo { NumberDecimalSeparator = ",", NumberGroupSeparator = "." };
                currentCell.Value = XLCellValue.FromObject(dataValue, numberFormat);
                break;
            case ColumnType.Date:
                var value = dataValue.ToString();
                currentCell.SetValue(string.IsNullOrEmpty(value) ? (DateTime?)null : DateTime.Parse(value));
                break;
            case ColumnType.Standard:
                currentCell.Value = XLCellValue.FromObject(dataValue);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(columnType));
        }
    }

    private static void SetCellType(ColumnType columnType, IXLCell currentCell)
    {
        switch (columnType)
        {
            case ColumnType.Accounting:
                currentCell.Style.NumberFormat.Format = AccountingFormat;
                break;
            case ColumnType.Date:
                currentCell.Style.DateFormat.SetFormat(DateFormat);
                break;
            case ColumnType.Standard:
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(columnType));
        }
    }

    private static DataTable GetData(string sql)
    {
        _configuration ??= new ConfigurationBuilder()
            .AddUserSecrets<ReportGenerator>()
            .Build();

        using var connection = new SqlConnection(_configuration.GetConnectionString(ConnectionName));
        connection.Open();

        using var command = new SqlCommand(sql, connection);
        using var adapter = new SqlDataAdapter(command);

        var dataTable = new DataTable();
        adapter.Fill(dataTable);
        return dataTable;
    }

    #endregion
}