using BomWizard.Services;

namespace BomWizard;

public sealed class Form1 : Form
{
    private readonly TextBox _excelPathTextBox = new();
    private readonly Button _browseButton = new();
    private readonly Button _updateButton = new();
    private readonly TextBox _logTextBox = new();
    private readonly Label _statusLabel = new();

    public Form1()
    {
        BuildInterface();
    }

    private void BuildInterface()
    {
        Text = "BOM Wizard - SolidWorks Balloon Sequence";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(760, 470);
        Size = new Size(900, 540);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(16),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var titleLabel = new Label
        {
            AutoSize = true,
            Font = new Font(Font.FontFamily, 14, FontStyle.Bold),
            Text = "Update SolidWorks BOM balloon sequence from Excel",
            Margin = new Padding(0, 0, 0, 12),
        };

        var pickerPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 3,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 12),
        };
        pickerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        pickerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        pickerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var pathLabel = new Label
        {
            AutoSize = true,
            Text = "Excel file:",
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 7, 8, 0),
        };

        _excelPathTextBox.Dock = DockStyle.Fill;
        _excelPathTextBox.PlaceholderText = "Select .xlsx, .xlsm, or .xls file with part numbers in column B and sequence in column C";

        _browseButton.Text = "Browse...";
        _browseButton.AutoSize = true;
        _browseButton.Margin = new Padding(8, 0, 0, 0);
        _browseButton.Click += BrowseButton_Click;

        pickerPanel.Controls.Add(pathLabel, 0, 0);
        pickerPanel.Controls.Add(_excelPathTextBox, 1, 0);
        pickerPanel.Controls.Add(_browseButton, 2, 0);

        _updateButton.Text = "Update Active SolidWorks BOM";
        _updateButton.AutoSize = true;
        _updateButton.Padding = new Padding(10, 4, 10, 4);
        _updateButton.Margin = new Padding(0, 0, 0, 12);
        _updateButton.Click += UpdateButton_Click;

        _logTextBox.Dock = DockStyle.Fill;
        _logTextBox.Multiline = true;
        _logTextBox.ReadOnly = true;
        _logTextBox.ScrollBars = ScrollBars.Vertical;
        _logTextBox.Font = new Font("Consolas", 9);

        _statusLabel.AutoSize = true;
        _statusLabel.Text = "Ready. Open the SolidWorks drawing or assembly that contains the BOM before updating.";
        _statusLabel.Margin = new Padding(0, 10, 0, 0);

        root.Controls.Add(titleLabel, 0, 0);
        root.Controls.Add(pickerPanel, 0, 1);
        root.Controls.Add(_updateButton, 0, 2);
        root.Controls.Add(_logTextBox, 0, 3);
        root.Controls.Add(_statusLabel, 0, 4);

        Controls.Add(root);
    }

    private void BrowseButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            Title = "Select BOM sequence Excel file",
            CheckFileExists = true,
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _excelPathTextBox.Text = dialog.FileName;
        }
    }

    private async void UpdateButton_Click(object? sender, EventArgs e)
    {
        var excelPath = _excelPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
        {
            MessageBox.Show(this, "Select a valid Excel file first.", "BOM Wizard", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SetBusy(true);
        Log("Reading Excel file...");

        try
        {
            var result = await Task.Run(() =>
            {
                var sequences = ExcelBomReader.ReadSequences(excelPath);
                var updateResult = SolidWorksBomUpdater.UpdateActiveBom(sequences);
                return (sequences.Count, updateResult);
            });

            Log($"Loaded {result.Count} Excel sequence row(s).");
            Log($"Scanned {result.updateResult.TablesScanned} BOM table(s).");
            Log($"Updated {result.updateResult.RowsUpdated} BOM row(s).");

            foreach (var message in result.updateResult.Messages)
            {
                Log(message);
            }

            _statusLabel.Text = result.updateResult.RowsUpdated > 0
                ? "Update complete. Review and save the SolidWorks document when ready."
                : "No BOM rows were updated. Check that part numbers in Excel match the active BOM.";
        }
        catch (Exception ex)
        {
            Log("ERROR: " + ex.Message);
            _statusLabel.Text = "Update failed.";
            MessageBox.Show(this, ex.Message, "BOM Wizard", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void SetBusy(bool busy)
    {
        _browseButton.Enabled = !busy;
        _updateButton.Enabled = !busy;
        _excelPathTextBox.Enabled = !busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
    }

    private void Log(string message)
    {
        _logTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
    }
}
