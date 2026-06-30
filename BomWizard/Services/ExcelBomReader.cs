using ClosedXML.Excel;

namespace BomWizard.Services;

public static class ExcelBomReader
{
    private static readonly string[] RequiredHeaders = ["P/N", "DESCRIPTION 1", "BSEQ", "OPSEQ"];

    public static IReadOnlyList<BomSequenceEntry> ReadSequences(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("The selected Excel file could not be found.", path);
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.First();
        var usedRange = worksheet.RangeUsed();

        if (usedRange is null)
        {
            throw new InvalidOperationException("The selected Excel worksheet is empty.");
        }

        var headerMap = FindHeaderMap(worksheet, usedRange);
        var headerRow = headerMap.HeaderRow;
        var entries = new List<BomSequenceEntry>();
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;

        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var partNumber = GetCell(worksheet, row, headerMap.Columns["P/N"]);
            var description1 = GetCell(worksheet, row, headerMap.Columns["DESCRIPTION 1"]);
            var bSeq = GetCell(worksheet, row, headerMap.Columns["BSEQ"]);
            var opSeq = GetCell(worksheet, row, headerMap.Columns["OPSEQ"]);

            if (string.IsNullOrWhiteSpace(partNumber)
                && string.IsNullOrWhiteSpace(description1)
                && string.IsNullOrWhiteSpace(bSeq)
                && string.IsNullOrWhiteSpace(opSeq))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(partNumber))
            {
                throw new InvalidOperationException($"Row {row} has BOM data but no P/N value.");
            }

            if (string.IsNullOrWhiteSpace(description1))
            {
                throw new InvalidOperationException($"Row {row} has P/N '{partNumber}' but no DESCRIPTION 1 value.");
            }

            if (string.IsNullOrWhiteSpace(bSeq))
            {
                throw new InvalidOperationException($"Row {row} has P/N '{partNumber}' but no BSEQ value.");
            }

            if (string.IsNullOrWhiteSpace(opSeq))
            {
                throw new InvalidOperationException($"Row {row} has P/N '{partNumber}' but no OPSEQ value.");
            }

            entries.Add(new BomSequenceEntry(partNumber, description1, bSeq, opSeq));
        }

        if (entries.Count == 0)
        {
            throw new InvalidOperationException("No BOM reference rows were found below the header row.");
        }

        return entries
            .OrderBy(entry => ParseSequence(entry.BSeq))
            .ThenBy(entry => entry.PartNumber, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static HeaderMap FindHeaderMap(IXLWorksheet worksheet, IXLRange usedRange)
    {
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = Math.Min(usedRange.RangeAddress.LastAddress.RowNumber, firstRow + 20);
        var firstColumn = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = usedRange.RangeAddress.LastAddress.ColumnNumber;

        for (var row = firstRow; row <= lastRow; row++)
        {
            var columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (var column = firstColumn; column <= lastColumn; column++)
            {
                var value = NormalizeHeader(worksheet.Cell(row, column).GetString());
                foreach (var requiredHeader in RequiredHeaders)
                {
                    if (value.Equals(NormalizeHeader(requiredHeader), StringComparison.OrdinalIgnoreCase))
                    {
                        columns[requiredHeader] = column;
                    }
                }
            }

            if (RequiredHeaders.All(columns.ContainsKey))
            {
                return new HeaderMap(row, columns);
            }
        }

        throw new InvalidOperationException("Could not find required headers: P/N, DESCRIPTION 1, BSEQ, OPSEQ.");
    }

    private static string GetCell(IXLWorksheet worksheet, int row, int column)
    {
        return worksheet.Cell(row, column).GetFormattedString().Trim();
    }

    private static string NormalizeHeader(string value)
    {
        return BomSequenceEntry.Normalize(value).Replace(".", string.Empty, StringComparison.Ordinal);
    }

    private static int ParseSequence(string value)
    {
        return int.TryParse(value, out var sequence) ? sequence : int.MaxValue;
    }

    private sealed record HeaderMap(int HeaderRow, Dictionary<string, int> Columns);
}
