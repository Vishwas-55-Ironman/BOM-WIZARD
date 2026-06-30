namespace BomWizard.Services;

public sealed record BomSequenceEntry(string PartNumber, string Description1, string BSeq, string OpSeq)
{
    public string CompositeKey => CreateCompositeKey(PartNumber, Description1);

    public static string CreateCompositeKey(string partNumber, string description1)
    {
        return $"{Normalize(partNumber)}|{Normalize(description1)}";
    }

    public static string Normalize(string value)
    {
        return string.Join(' ', value.Trim().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }
}
