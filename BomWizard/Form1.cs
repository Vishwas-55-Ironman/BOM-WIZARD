using System.Drawing.Drawing2D;
using BomWizard.Services;

namespace BomWizard;

public sealed class Form1 : Form
{
    private readonly TextBox _excelPathTextBox = new();
    private readonly GradientButton _browseButton = new("Browse");
    private readonly GradientButton _updateButton = new("Update Active SolidWorks BOM");
    private readonly ComboBox _matchModeComboBox = new();
    private readonly CheckBox _updateBSeqCheckBox = new();
    private readonly CheckBox _updateOpSeqCheckBox = new();
    private readonly TextBox _logTextBox = new();
    private readonly Label _statusLabel = new();

    public Form1()
    {
        BuildInterface();
    }

    private void BuildInterface()
    {
        Text = "BOM Wizard - SolidWorks BOM Sort";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(860, 560);
        Size = new Size(1020, 640);
        Font = new Font("Segoe UI", 10F);

        var shell = new GradientPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24),
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            BackColor = Color.Transparent,
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var titleLabel = new Label
        {
            AutoSize = true,
            BackColor = Color.Transparent,
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 20F, FontStyle.Bold),
            Text = "SolidWorks BOM sorter",
            Margin = new Padding(0, 0, 0, 22),
        };

        var pickerPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 3,
            AutoSize = true,
            BackColor = Color.Transparent,
            Margin = new Padding(0, 0, 0, 16),
        };
        pickerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        pickerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        pickerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var pathLabel = new Label
        {
            AutoSize = true,
            Text = "Reference Excel",
            Anchor = AnchorStyles.Left,
            BackColor = Color.Transparent,
            ForeColor = Color.White,
            Margin = new Padding(0, 8, 12, 0),
        };

        _excelPathTextBox.Dock = DockStyle.Fill;
        _excelPathTextBox.PlaceholderText = "Select workbook with headers P/N, DESCRIPTION 1, BSEQ, OPSEQ";
        _excelPathTextBox.BorderStyle = BorderStyle.FixedSingle;
        _excelPathTextBox.Margin = new Padding(0, 0, 10, 0);

        _browseButton.AutoSize = false;
        _browseButton.Size = new Size(118, 36);
        _browseButton.Click += BrowseButton_Click;

        pickerPanel.Controls.Add(pathLabel, 0, 0);
        pickerPanel.Controls.Add(_excelPathTextBox, 1, 0);
        pickerPanel.Controls.Add(_browseButton, 2, 0);

        _updateButton.AutoSize = false;
        _updateButton.Size = new Size(270, 44);
        _updateButton.Margin = new Padding(0, 0, 0, 16);
        _updateButton.Click += UpdateButton_Click;

        var optionsPanel = BuildOptionsPanel();

        _logTextBox.Dock = DockStyle.Fill;
        _logTextBox.Multiline = true;
        _logTextBox.ReadOnly = true;
        _logTextBox.ScrollBars = ScrollBars.Vertical;
        _logTextBox.Font = new Font("Cascadia Mono", 9F);
        _logTextBox.BackColor = Color.FromArgb(246, 249, 253);
        _logTextBox.ForeColor = Color.FromArgb(25, 36, 58);
        _logTextBox.BorderStyle = BorderStyle.FixedSingle;

        _statusLabel.AutoSize = true;
        _statusLabel.Text = "Ready.";
        _statusLabel.BackColor = Color.Transparent;
        _statusLabel.ForeColor = Color.White;
        _statusLabel.Margin = new Padding(0, 14, 0, 0);

        root.Controls.Add(titleLabel, 0, 0);
        root.Controls.Add(pickerPanel, 0, 1);
        root.Controls.Add(optionsPanel, 0, 2);
        root.Controls.Add(_updateButton, 0, 3);
        root.Controls.Add(_logTextBox, 0, 4);
        root.Controls.Add(_statusLabel, 0, 5);

        shell.Controls.Add(root);
        Controls.Add(shell);
    }

    private Control BuildOptionsPanel()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 90,
            BackColor = Color.FromArgb(245, 248, 255),
            Margin = new Padding(0, 0, 0, 16),
            Padding = new Padding(14, 12, 14, 12),
        };

        var title = new Label
        {
            AutoSize = true,
            Text = "Options",
            Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
            ForeColor = Color.FromArgb(28, 39, 64),
            Location = new Point(14, 10),
        };

        var matchLabel = new Label
        {
            AutoSize = true,
            Text = "Sort / match by",
            ForeColor = Color.FromArgb(28, 39, 64),
            Location = new Point(14, 42),
        };

        _matchModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        _matchModeComboBox.Location = new Point(118, 38);
        _matchModeComboBox.Size = new Size(315, 28);
        _matchModeComboBox.Items.Add(new MatchModeItem("Auto: P/N + DESCRIPTION 1, then DESCRIPTION 1", BomMatchMode.PartAndDescriptionThenDescription));
        _matchModeComboBox.Items.Add(new MatchModeItem("P/N + DESCRIPTION 1", BomMatchMode.PartAndDescription));
        _matchModeComboBox.Items.Add(new MatchModeItem("P/N only", BomMatchMode.PartNumber));
        _matchModeComboBox.Items.Add(new MatchModeItem("DESCRIPTION 1 only", BomMatchMode.Description));
        _matchModeComboBox.SelectedIndex = 0;

        var flow = new FlowLayoutPanel
        {
            AutoSize = false,
            Location = new Point(460, 40),
            Size = new Size(360, 34),
            BackColor = Color.Transparent,
            WrapContents = false,
        };

        ConfigureOption(_updateBSeqCheckBox, "Update BSEQ", isChecked: true, isEnabled: true);
        ConfigureOption(_updateOpSeqCheckBox, "Update OPSEQ", isChecked: true, isEnabled: true);

        flow.Controls.Add(_updateBSeqCheckBox);
        flow.Controls.Add(_updateOpSeqCheckBox);

        panel.Controls.Add(title);
        panel.Controls.Add(matchLabel);
        panel.Controls.Add(_matchModeComboBox);
        panel.Controls.Add(flow);
        return panel;
    }

    private static void ConfigureOption(CheckBox checkBox, string text, bool isChecked, bool isEnabled)
    {
        checkBox.Text = text;
        checkBox.Checked = isChecked;
        checkBox.Enabled = isEnabled;
        checkBox.AutoSize = true;
        checkBox.Margin = new Padding(0, 0, 26, 0);
        checkBox.ForeColor = Color.FromArgb(28, 39, 64);
        checkBox.FlatStyle = FlatStyle.System;
    }

    private void BrowseButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            Title = "Select BOM reference Excel file",
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
            MessageBox.Show(this, "Select a valid Excel reference file first.", "BOM Wizard", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SetBusy(true);
        Log("Reading reference workbook...");

        try
        {
            var options = new BomUpdateOptions(
                GetSelectedMatchMode(),
                _updateBSeqCheckBox.Checked,
                _updateOpSeqCheckBox.Checked);

            var result = await Task.Run(() =>
            {
                var rows = ExcelBomReader.ReadSequences(excelPath);
                var updateResult = SolidWorksBomUpdater.UpdateActiveBom(rows, options);
                return (rows.Count, updateResult);
            });

            Log($"Loaded {result.Count} reference row(s).");
            Log($"Scanned {result.updateResult.TablesScanned} BOM table(s).");
            Log($"Updated {result.updateResult.RowsUpdated} BOM row(s), {result.updateResult.CellsUpdated} cell(s).");

            foreach (var message in result.updateResult.Messages)
            {
                Log(message);
            }

            _statusLabel.Text = result.updateResult.RowsUpdated > 0
                ? "Update complete. Review the SolidWorks BOM, then save the document when ready."
                : "No BOM rows were updated. Check P/N and DESCRIPTION 1 in the active BOM.";
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
        _matchModeComboBox.Enabled = !busy;
        _updateBSeqCheckBox.Enabled = !busy;
        _updateOpSeqCheckBox.Enabled = !busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
    }

    private void Log(string message)
    {
        _logTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
    }

    private BomMatchMode GetSelectedMatchMode()
    {
        return _matchModeComboBox.SelectedItem is MatchModeItem item
            ? item.MatchMode
            : BomMatchMode.PartAndDescriptionThenDescription;
    }

    private sealed record MatchModeItem(string Text, BomMatchMode MatchMode)
    {
        public override string ToString()
        {
            return Text;
        }
    }
}

internal sealed class GradientPanel : Panel
{
    protected override void OnPaintBackground(PaintEventArgs e)
    {
        using var brush = new LinearGradientBrush(
            ClientRectangle,
            Color.FromArgb(30, 58, 138),
            Color.FromArgb(190, 24, 93),
            LinearGradientMode.ForwardDiagonal);

        e.Graphics.FillRectangle(brush, ClientRectangle);

        using var overlay = new SolidBrush(Color.FromArgb(72, 6, 15, 34));
        e.Graphics.FillRectangle(overlay, ClientRectangle);
    }
}

internal sealed class GradientButton : Button
{
    private bool _hovered;

    public GradientButton(string text)
    {
        Text = text;
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        BackColor = Color.FromArgb(255, 255, 255);
        ForeColor = Color.FromArgb(28, 39, 64);
        Font = new Font("Segoe UI Semibold", 9.8F, FontStyle.Bold);
        Cursor = Cursors.Hand;
        UseVisualStyleBackColor = false;
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        _hovered = true;
        Invalidate();
        base.OnMouseEnter(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        _hovered = false;
        Invalidate();
        base.OnMouseLeave(e);
    }

    protected override void OnPaint(PaintEventArgs pevent)
    {
        pevent.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var bounds = new Rectangle(0, 0, Width - 1, Height - 1);

        using var fillBrush = new SolidBrush(Enabled ? BackColor : Color.FromArgb(225, 229, 235));
        pevent.Graphics.FillRectangle(fillBrush, bounds);

        if (_hovered && Enabled)
        {
            using var borderBrush = new LinearGradientBrush(
                bounds,
                Color.FromArgb(0, 122, 255),
                Color.FromArgb(236, 72, 153),
                LinearGradientMode.Horizontal);
            using var pen = new Pen(borderBrush, 3F);
            pevent.Graphics.DrawRectangle(pen, bounds);
        }
        else
        {
            using var pen = new Pen(Color.FromArgb(210, 218, 232), 1F);
            pevent.Graphics.DrawRectangle(pen, bounds);
        }

        TextRenderer.DrawText(
            pevent.Graphics,
            Text,
            Font,
            bounds,
            Enabled ? ForeColor : Color.FromArgb(130, 138, 150),
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }
}
