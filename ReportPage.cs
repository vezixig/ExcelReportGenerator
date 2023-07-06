namespace ReportGenerator;

/// <summary>Represents a report page.</summary>
internal class ReportPage
{
    #region Properties

    /// <summary>Gets or sets the dictionary of special columns and their corresponding types.</summary>
    public Dictionary<string, ColumnType> SpecialColumns { get; set; } = new();

    /// <summary>Gets or sets the name of the report page.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the SQL query associated with the report page.</summary>
    public string Sql { get; set; } = string.Empty;

    #endregion
}