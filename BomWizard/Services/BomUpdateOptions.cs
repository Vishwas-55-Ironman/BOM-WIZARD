namespace BomWizard.Services;

public enum BomMatchMode
{
    PartAndDescriptionThenDescription,
    PartAndDescription,
    PartNumber,
    Description
}

public sealed record BomUpdateOptions(BomMatchMode MatchMode, bool UpdateBSeq, bool UpdateOpSeq)
{
    public static BomUpdateOptions Default { get; } = new(
        BomMatchMode.PartAndDescriptionThenDescription,
        true,
        true);
}
