namespace BomWizard.Services;

public sealed class BomUpdateResult
{
    public int TablesScanned { get; set; }
    public int RowsUpdated { get; set; }
    public List<string> Messages { get; } = [];
}
