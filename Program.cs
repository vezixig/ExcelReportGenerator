using ReportGenerator;

const string sql = """
    SELECT 
        Date,
        Sold,
        Price,
        Seller
    FROM invoices
    """;

var specialColumns = new Dictionary<string, ColumnType>
{
    { "Price", ColumnType.Accounting },
    { "Date", ColumnType.Date }
};

var report = new ReportDefinition
{
    FilePath = "D:\\MyReport.xlsx",
    Pages = new()
    {
        new()
        {
            Sql = sql,
            SpecialColumns = specialColumns,
            Name = "Invoices"
        }
    }
};

var generator = new ReportGenerator.ReportGenerator();
generator.GenerateReport(new() { report });