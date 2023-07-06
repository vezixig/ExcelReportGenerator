namespace ReportGenerator;

/// <summary>Represents a report definition.</summary>
internal class ReportDefinition
{
    #region Properties

    /// <summary>Gets or sets the filename the report is saved to.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Gets or sets the list of report pages associated with the report definition.</summary>
    public List<ReportPage> Pages { get; set; } = new();

    #endregion
}