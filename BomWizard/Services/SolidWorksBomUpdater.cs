using System.Reflection;
using System.Runtime.InteropServices;

namespace BomWizard.Services;

public static class SolidWorksBomUpdater
{
    private static readonly string[] PartHeaders = ["P/N", "PN", "PART NUMBER", "PART NO", "PART NO.", "PARTNUMBER"];
    private static readonly string[] DescriptionHeaders = ["DESCRIPTION 1", "DESCRIPTION", "DESC 1", "DESC"];
    private static readonly string[] BSeqHeaders = ["BSEQ", "BALLOON SEQUENCE", "BALLOON", "ITEM", "ITEM NO", "ITEM NO.", "SEQ"];
    private static readonly string[] OpSeqHeaders = ["OPSEQ", "OP SEQ", "OPERATION SEQUENCE", "OPERATION"];

    public static BomUpdateResult UpdateActiveBom(IReadOnlyList<BomSequenceEntry> entries)
    {
        return UpdateActiveBom(entries, BomUpdateOptions.Default);
    }

    public static BomUpdateResult UpdateActiveBom(IReadOnlyList<BomSequenceEntry> entries, BomUpdateOptions options)
    {
        if (entries.Count == 0)
        {
            throw new InvalidOperationException("No Excel reference rows were loaded.");
        }

        if (!options.UpdateBSeq && !options.UpdateOpSeq)
        {
            throw new InvalidOperationException("Select at least one value to update: BSEQ or OPSEQ.");
        }

        var reference = BomReference.Create(entries);
        var result = new BomUpdateResult();
        var matchedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var swApp = ActiveComObject.Get("SldWorks.Application");
        var activeDoc = ComDispatch.GetProperty(swApp, "ActiveDoc");

        if (activeDoc is null)
        {
            throw new InvalidOperationException("SolidWorks is running, but no document is active.");
        }

        foreach (var bomFeature in EnumerateBomFeatures(activeDoc))
        {
            foreach (var table in GetTableAnnotations(bomFeature))
            {
                result.TablesScanned++;
                UpdateTable(table, reference, options, matchedKeys, result);
            }
        }

        if (result.TablesScanned == 0)
        {
            result.Messages.Add("No BOM tables were found in the active SolidWorks document.");
        }

        foreach (var missing in entries.Where(entry => !matchedKeys.Contains(entry.CompositeKey)).OrderBy(entry => entry.BSeq))
        {
            result.Messages.Add($"No matching BOM row found for P/N '{missing.PartNumber}' / DESCRIPTION 1 '{missing.Description1}'.");
        }

        TryInvoke(activeDoc, "ForceRebuild3", false);
        return result;
    }

    private static void UpdateTable(
        object table,
        BomReference reference,
        BomUpdateOptions options,
        HashSet<string> matchedKeys,
        BomUpdateResult result)
    {
        var rowCount = ComDispatch.GetInt(table, "RowCount");
        var columnCount = ComDispatch.GetInt(table, "ColumnCount");
        var columns = FindBomColumns(table, rowCount, columnCount);

        if (RequiresPartNumber(options.MatchMode) && columns.PartNumber is null)
        {
            result.Messages.Add("Skipped a BOM table because no P/N or part number column was found.");
            return;
        }

        if (RequiresDescription(options.MatchMode) && columns.Description1 is null)
        {
            result.Messages.Add("Skipped a BOM table because no DESCRIPTION 1 column was found.");
            return;
        }

        if (options.UpdateBSeq && columns.BSeq is null)
        {
            result.Messages.Add("A matching BOM table has no BSEQ/balloon/item column, so BSEQ values cannot be written.");
        }

        if (options.UpdateOpSeq && columns.OpSeq is null)
        {
            result.Messages.Add("A matching BOM table has no OPSEQ/operation sequence column, so OPSEQ values cannot be written.");
        }

        for (var row = columns.FirstDataRow; row < rowCount; row++)
        {
            var entry = FindMatchingEntry(table, row, columnCount, columns, reference, options);
            if (entry is null)
            {
                continue;
            }

            var changedCells = 0;
            if (options.UpdateBSeq && columns.BSeq is int bSeqColumn)
            {
                SetCellText(table, row, bSeqColumn, entry.BSeq);
                changedCells++;
            }

            if (options.UpdateOpSeq && columns.OpSeq is int opSeqColumn)
            {
                SetCellText(table, row, opSeqColumn, entry.OpSeq);
                changedCells++;
            }

            if (changedCells == 0)
            {
                continue;
            }

            matchedKeys.Add(entry.CompositeKey);
            result.RowsUpdated++;
            result.CellsUpdated += changedCells;
            result.Messages.Add($"Updated row {row + 1}: P/N '{entry.PartNumber}', BSEQ '{entry.BSeq}', OPSEQ '{entry.OpSeq}'.");
        }

        TryInvoke(table, "UpdateTableAnnotation");
    }

    private static BomColumns FindBomColumns(object table, int rowCount, int columnCount)
    {
        var headerRowsToCheck = Math.Min(rowCount, 5);

        for (var row = 0; row < headerRowsToCheck; row++)
        {
            int? partNumber = null;
            int? description1 = null;
            int? bSeq = null;
            int? opSeq = null;

            for (var column = 0; column < columnCount; column++)
            {
                var header = NormalizeHeader(GetCellText(table, row, column));

                partNumber ??= MatchesAny(header, PartHeaders) ? column : null;
                description1 ??= MatchesAny(header, DescriptionHeaders) ? column : null;
                bSeq ??= MatchesAny(header, BSeqHeaders) ? column : null;
                opSeq ??= MatchesAny(header, OpSeqHeaders) ? column : null;
            }

            if (partNumber is not null)
            {
                return new BomColumns(partNumber, description1, bSeq, opSeq, row + 1);
            }
        }

        return new BomColumns(null, null, null, null, 1);
    }

    private static BomSequenceEntry? FindMatchingEntry(
        object table,
        int row,
        int columnCount,
        BomColumns columns,
        BomReference reference,
        BomUpdateOptions options)
    {
        var partNumber = columns.PartNumber is int partColumn
            ? GetComparableCell(table, row, partColumn)
            : string.Empty;
        var description1 = columns.Description1 is int descriptionColumn
            ? GetComparableCell(table, row, descriptionColumn)
            : string.Empty;

        if (options.MatchMode is BomMatchMode.PartAndDescription or BomMatchMode.PartAndDescriptionThenDescription
            && !string.IsNullOrWhiteSpace(partNumber)
            && !string.IsNullOrWhiteSpace(description1))
        {
            var compositeKey = BomSequenceEntry.CreateCompositeKey(partNumber, description1);
            if (reference.ByComposite.TryGetValue(compositeKey, out var entry))
            {
                return entry;
            }
        }

        if (options.MatchMode is BomMatchMode.PartNumber
            && !string.IsNullOrWhiteSpace(partNumber)
            && reference.UniqueByPartNumber.TryGetValue(BomSequenceEntry.Normalize(partNumber), out var partOnlyEntry))
        {
            return partOnlyEntry;
        }

        if (options.MatchMode is BomMatchMode.Description or BomMatchMode.PartAndDescriptionThenDescription
            && !string.IsNullOrWhiteSpace(description1)
            && reference.UniqueByDescription1.TryGetValue(BomSequenceEntry.Normalize(description1), out var descriptionEntry))
        {
            return descriptionEntry;
        }

        for (var column = 0; column < columnCount; column++)
        {
            var cellText = GetComparableCell(table, row, column);
            if (string.IsNullOrWhiteSpace(cellText))
            {
                continue;
            }

            if (options.MatchMode is BomMatchMode.PartNumber
                && reference.UniqueByPartNumber.TryGetValue(BomSequenceEntry.Normalize(cellText), out var scannedEntry))
            {
                return scannedEntry;
            }

            var fileStem = Path.GetFileNameWithoutExtension(cellText);
            if (options.MatchMode is BomMatchMode.PartNumber
                && !string.IsNullOrWhiteSpace(fileStem)
                && reference.UniqueByPartNumber.TryGetValue(BomSequenceEntry.Normalize(fileStem), out scannedEntry))
            {
                return scannedEntry;
            }
        }

        return null;
    }

    private static IEnumerable<object> EnumerateBomFeatures(object modelDoc)
    {
        var firstFeature = ComDispatch.Invoke(modelDoc, "FirstFeature");

        for (var feature = firstFeature; feature is not null; feature = TryInvoke(feature, "GetNextFeature"))
        {
            foreach (var bomFeature in EnumerateBomFeaturesFromFeature(feature))
            {
                yield return bomFeature;
            }
        }
    }

    private static IEnumerable<object> EnumerateBomFeaturesFromFeature(object feature)
    {
        var typeName = Convert.ToString(TryInvoke(feature, "GetTypeName2")) ?? string.Empty;
        if (typeName.Equals("BomFeat", StringComparison.OrdinalIgnoreCase))
        {
            var bomFeature = TryInvoke(feature, "GetSpecificFeature2");
            if (bomFeature is not null)
            {
                yield return bomFeature;
            }
        }

        var firstSubFeature = TryInvoke(feature, "GetFirstSubFeature");
        for (var subFeature = firstSubFeature; subFeature is not null; subFeature = TryInvoke(subFeature, "GetNextSubFeature"))
        {
            foreach (var bomFeature in EnumerateBomFeaturesFromFeature(subFeature))
            {
                yield return bomFeature;
            }
        }
    }

    private static IEnumerable<object> GetTableAnnotations(object bomFeature)
    {
        var annotations = TryInvoke(bomFeature, "GetTableAnnotations");

        if (annotations is null)
        {
            yield break;
        }

        if (annotations is Array array)
        {
            foreach (var annotation in array)
            {
                if (annotation is not null)
                {
                    yield return annotation;
                }
            }

            yield break;
        }

        yield return annotations;
    }

    private static string GetComparableCell(object table, int row, int column)
    {
        return BomSequenceEntry.Normalize(GetCellText(table, row, column));
    }

    private static string GetCellText(object table, int row, int column)
    {
        try
        {
            return ComDispatch.GetString(table, "Text", row, column);
        }
        catch (Exception ex) when (IsComDispatchException(ex))
        {
        }

        try
        {
            return ComDispatch.GetString(table, "DisplayedText", row, column);
        }
        catch (Exception ex) when (IsComDispatchException(ex))
        {
            return Convert.ToString(TryInvoke(table, "DisplayedText", row, column))?.Trim() ?? string.Empty;
        }
    }

    private static void SetCellText(object table, int row, int column, string value)
    {
        try
        {
            ComDispatch.SetProperty(table, "Text", row, column, value);
            return;
        }
        catch (Exception ex) when (IsComDispatchException(ex))
        {
        }

        if (!TryInvoke(table, "SetCellText", out _, row, column, value))
        {
            throw new InvalidOperationException($"Could not update BOM cell at row {row + 1}, column {column + 1}.");
        }
    }

    private static string NormalizeHeader(string value)
    {
        return BomSequenceEntry.Normalize(value)
            .Replace(".", string.Empty, StringComparison.Ordinal)
            .Replace("/", string.Empty, StringComparison.Ordinal)
            .ToUpperInvariant();
    }

    private static bool MatchesAny(string normalizedHeader, IEnumerable<string> candidates)
    {
        return candidates
            .Select(NormalizeHeader)
            .Any(candidate => normalizedHeader.Equals(candidate, StringComparison.OrdinalIgnoreCase));
    }

    private static bool RequiresPartNumber(BomMatchMode matchMode)
    {
        return matchMode is BomMatchMode.PartAndDescription
            or BomMatchMode.PartNumber;
    }

    private static bool RequiresDescription(BomMatchMode matchMode)
    {
        return matchMode is BomMatchMode.PartAndDescription
            or BomMatchMode.Description
            or BomMatchMode.PartAndDescriptionThenDescription;
    }

    private static object? TryInvoke(object target, string methodName, params object[] args)
    {
        return TryInvoke(target, methodName, out var value, args) ? value : null;
    }

    private static bool TryInvoke(object target, string methodName, out object? value, params object[] args)
    {
        try
        {
            value = ComDispatch.Invoke(target, methodName, args);
            return true;
        }
        catch (Exception ex) when (IsComDispatchException(ex))
        {
            value = null;
            return false;
        }
    }

    private static bool IsComDispatchException(Exception ex)
    {
        return ex is COMException or MissingMethodException
            || ex is TargetInvocationException { InnerException: COMException };
    }

    private sealed record BomColumns(int? PartNumber, int? Description1, int? BSeq, int? OpSeq, int FirstDataRow);

    private sealed class BomReference
    {
        public Dictionary<string, BomSequenceEntry> ByComposite { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, BomSequenceEntry> UniqueByPartNumber { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, BomSequenceEntry> UniqueByDescription1 { get; } = new(StringComparer.OrdinalIgnoreCase);

        public static BomReference Create(IReadOnlyList<BomSequenceEntry> entries)
        {
            var reference = new BomReference();
            foreach (var entry in entries)
            {
                reference.ByComposite[entry.CompositeKey] = entry;
            }

            foreach (var group in entries.GroupBy(entry => BomSequenceEntry.Normalize(entry.PartNumber), StringComparer.OrdinalIgnoreCase))
            {
                if (group.Count() == 1)
                {
                    reference.UniqueByPartNumber[group.Key] = group.Single();
                }
            }

            foreach (var group in entries.GroupBy(entry => BomSequenceEntry.Normalize(entry.Description1), StringComparer.OrdinalIgnoreCase))
            {
                if (group.Count() == 1)
                {
                    reference.UniqueByDescription1[group.Key] = group.Single();
                }
            }

            return reference;
        }
    }
}
