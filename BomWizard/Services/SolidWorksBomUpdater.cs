using System.Reflection;
using System.Runtime.InteropServices;

namespace BomWizard.Services;

public static class SolidWorksBomUpdater
{
    private static readonly string[] ItemHeaderTerms = ["item", "item no", "item no.", "balloon", "balloon sequence", "seq"];

    public static BomUpdateResult UpdateActiveBom(IReadOnlyDictionary<string, BomSequenceEntry> sequences)
    {
        if (sequences.Count == 0)
        {
            throw new InvalidOperationException("No Excel sequences were loaded.");
        }

        var result = new BomUpdateResult();
        var matchedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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
                UpdateTable(table, sequences, matchedParts, result);
            }
        }

        if (result.TablesScanned == 0)
        {
            result.Messages.Add("No BOM tables were found in the active SolidWorks document.");
        }

        foreach (var missing in sequences.Keys.Where(key => !matchedParts.Contains(key)).OrderBy(key => key))
        {
            result.Messages.Add($"No matching BOM row found for Excel part number '{sequences[missing].PartNumber}'.");
        }

        TryInvoke(activeDoc, "ForceRebuild3", false);
        return result;
    }

    private static void UpdateTable(
        object table,
        IReadOnlyDictionary<string, BomSequenceEntry> sequences,
        HashSet<string> matchedParts,
        BomUpdateResult result)
    {
        var rowCount = ComDispatch.GetInt(table, "RowCount");
        var columnCount = ComDispatch.GetInt(table, "ColumnCount");
        var itemColumn = FindItemColumn(table, rowCount, columnCount);

        for (var row = 1; row < rowCount; row++)
        {
            var matchKey = FindMatchingPartKey(table, row, columnCount, sequences);
            if (matchKey is null)
            {
                continue;
            }

            var sequence = sequences[matchKey].Sequence;
            SetCellText(table, row, itemColumn, sequence);
            matchedParts.Add(matchKey);
            result.RowsUpdated++;
            result.Messages.Add($"Updated row {row + 1}: part '{sequences[matchKey].PartNumber}' -> balloon sequence '{sequence}'.");
        }

        TryInvoke(table, "UpdateTableAnnotation");
    }

    private static int FindItemColumn(object table, int rowCount, int columnCount)
    {
        var headerRowsToCheck = Math.Min(rowCount, 3);

        for (var row = 0; row < headerRowsToCheck; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var header = GetCellText(table, row, column);
                if (ItemHeaderTerms.Any(term => header.Contains(term, StringComparison.OrdinalIgnoreCase)))
                {
                    return column;
                }
            }
        }

        return 0;
    }

    private static string? FindMatchingPartKey(
        object table,
        int row,
        int columnCount,
        IReadOnlyDictionary<string, BomSequenceEntry> sequences)
    {
        for (var column = 0; column < columnCount; column++)
        {
            var cellText = GetCellText(table, row, column);
            var normalized = ExcelBomReader.Normalize(cellText);

            if (sequences.ContainsKey(normalized))
            {
                return normalized;
            }

            var fileStem = Path.GetFileNameWithoutExtension(normalized);
            if (!string.IsNullOrWhiteSpace(fileStem) && sequences.ContainsKey(fileStem))
            {
                return fileStem;
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
}
