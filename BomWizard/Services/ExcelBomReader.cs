using ClosedXML.Excel;

namespace BomWizard.Services;

public static class ExcelBomReader
{
    public static IReadOnlyDictionary<string, BomSequenceEntry> ReadSequences(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("The selected Excel file could not be found.", path);
        }

        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheets.First();
        var usedRange = worksheet.RangeUsed();

        if (usedRange is null)
        {
            throw new InvalidOperationException("The selected Excel worksheet is empty.");
        }

        var entries = new Dictionary<string, BomSequenceEntry>(StringComparer.OrdinalIgnoreCase);
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;

        for (var row = firstRow; row <= lastRow; row++)
        {
            var partNumber = worksheet.Cell(row, 2).GetString().Trim();
            var sequence = worksheet.Cell(row, 3).GetFormattedString().Trim();

            if (string.IsNullOrWhiteSpace(partNumber) && string.IsNullOrWhiteSpace(sequence))
            {
                continue;
            }

            if (LooksLikeHeader(partNumber, sequence))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(partNumber))
            {
                throw new InvalidOperationException($"Row {row} has a sequence value but no part number in column B.");
            }

            if (string.IsNullOrWhiteSpace(sequence))
            {
                throw new InvalidOperationException($"Row {row} has part number '{partNumber}' but no sequence value in column C.");
            }

            entries[Normalize(partNumber)] = new BomSequenceEntry(partNumber, sequence);
        }

        if (entries.Count == 0)
        {
            throw new InvalidOperationException("No part number and sequence pairs were found in columns B and C.");
        }

        return entries;
    }

    internal static string Normalize(string value)
    {
        return value.Trim();
    }

    private static bool LooksLikeHeader(string partNumber, string sequence)
    {
        return partNumber.Contains("part", StringComparison.OrdinalIgnoreCase)
            && (sequence.Contains("sequence", StringComparison.OrdinalIgnoreCase)
                || sequence.Contains("balloon", StringComparison.OrdinalIgnoreCase)
                || sequence.Contains("item", StringComparison.OrdinalIgnoreCase));
    }
}
