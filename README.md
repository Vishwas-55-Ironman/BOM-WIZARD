# BOM Wizard

WinForms utility for updating SolidWorks BOM balloon/item sequence values from an Excel workbook.

## Workflow

1. Open the SolidWorks drawing or assembly that contains the BOM you want to update.
2. Run `BomWizard` from Visual Studio.
3. Browse to an Excel workbook (`.xlsx` or `.xlsm`).
4. Put part numbers in column B and the desired balloon sequence values in column C.
5. Click **Update Active SolidWorks BOM**.

## Matching behavior

- The app connects to the active running SolidWorks session through `SldWorks.Application`.
- It scans BOM features in the active document and reads their table annotations.
- For each BOM row, it looks for an Excel part number in any cell in that row.
- It updates the BOM item/balloon column, detected from headers such as `Item`, `Item No.`, `Balloon`, or `Sequence`.
- If no item/balloon header is found, it falls back to the first BOM column.

## Notes

- The workbook reader uses ClosedXML, so modern Excel formats (`.xlsx`, `.xlsm`) are supported.
- The project uses late-bound SolidWorks COM calls, so it can build without SolidWorks interop assemblies installed.
- Review the SolidWorks document after updating, then save it manually when the BOM looks correct.
