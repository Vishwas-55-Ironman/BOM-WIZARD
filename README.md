# BOM Wizard

WinForms utility for sorting/updating SolidWorks BOM sequence fields from a reference Excel workbook.

## Workflow

1. Open the SolidWorks drawing or assembly that contains the BOM you want to update.
2. Run `BomWizard` from Visual Studio.
3. Browse to an Excel workbook (`.xlsx` or `.xlsm`).
4. Use a reference sheet with headers `P/N`, `DESCRIPTION 1`, `BSEQ`, and `OPSEQ`.
5. Click **Update Active SolidWorks BOM**.

## Matching behavior

- The app connects to the active running SolidWorks session through `SldWorks.Application`.
- It scans BOM features in the active document and reads their table annotations.
- The **Sort / match by** dropdown controls row matching:
  - `Auto: P/N + DESCRIPTION 1, then DESCRIPTION 1`
  - `P/N + DESCRIPTION 1`
  - `P/N only`
  - `DESCRIPTION 1 only`
- Use `DESCRIPTION 1 only` when SolidWorks part numbers differ from the Excel reference but descriptions match.
- It writes `BSEQ` into a BOM column named `BSEQ`, `Balloon Sequence`, `Balloon`, `Item`, or similar.
- It writes `OPSEQ` into a BOM column named `OPSEQ`, `OP SEQ`, `Operation Sequence`, or similar.

## Notes

- The workbook reader uses ClosedXML, so modern Excel formats (`.xlsx`, `.xlsm`) are supported.
- The Excel file can be open in Excel while the app reads it.
- The project uses late-bound SolidWorks COM calls, so it can build without SolidWorks interop assemblies installed.
- Review the SolidWorks document after updating, then save it manually when the BOM looks correct.
