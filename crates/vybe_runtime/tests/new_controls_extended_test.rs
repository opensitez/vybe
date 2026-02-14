use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_new_controls_extended() {
    let source = std::fs::read_to_string("../../tests/test_new_controls_extended.vb")
        .expect("Failed to read test file");
    let program = parse_program(&source).expect("Parse error");

    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");

    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

    let output = interp.side_effects.iter().filter_map(|e| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = e {
            Some(msg.as_str())
        } else {
            None
        }
    }).collect::<Vec<_>>().join("\n");

    println!("Output:\n{}", output);

    let pass_count = output.matches("PASS:").count();
    let fail_count = output.matches("FAIL:").count();
    println!("Passes: {}, Failures: {}", pass_count, fail_count);

    // ===== CheckedListBox =====
    assert!(output.contains("PASS: CheckedListBox.Items not Nothing"), "CheckedListBox Items");
    assert!(output.contains("PASS: CheckedListBox.SelectedIndex default is -1"), "CheckedListBox SelectedIndex");
    assert!(output.contains("PASS: CheckedListBox.SelectionMode default is 1"), "CheckedListBox SelectionMode");
    assert!(output.contains("PASS: CheckedListBox.Sorted default is False"), "CheckedListBox Sorted default");
    assert!(output.contains("PASS: CheckedListBox.CheckOnClick default is False"), "CheckedListBox CheckOnClick default");
    assert!(output.contains("PASS: CheckedListBox.Items.Count after Add is 3"), "CheckedListBox Items.Count");
    assert!(output.contains("PASS: CheckedListBox.GetItemChecked(0) after SetItemChecked(0,True)"), "CheckedListBox SetItemChecked/GetItemChecked");
    assert!(output.contains("PASS: CheckedListBox.GetItemChecked(1) default is False"), "CheckedListBox GetItemChecked default");
    assert!(output.contains("PASS: CheckedListBox.GetItemChecked(1) after uncheck"), "CheckedListBox uncheck");
    assert!(output.contains("PASS: CheckedListBox.GetItemCheckState(0) is 1 (Checked)"), "CheckedListBox GetItemCheckState checked");
    assert!(output.contains("PASS: CheckedListBox.GetItemCheckState(1) is 0 (Unchecked)"), "CheckedListBox GetItemCheckState unchecked");
    assert!(output.contains("PASS: CheckedListBox.SetItemCheckState(2,1) works"), "CheckedListBox SetItemCheckState");
    assert!(output.contains("PASS: CheckedListBox.Sorted set to True"), "CheckedListBox Sorted set");
    assert!(output.contains("PASS: CheckedListBox.CheckOnClick set to True"), "CheckedListBox CheckOnClick set");

    // ===== DomainUpDown =====
    assert!(output.contains("PASS: DomainUpDown.Items not Nothing"), "DomainUpDown Items");
    assert!(output.contains("PASS: DomainUpDown.SelectedIndex default is -1"), "DomainUpDown SelectedIndex default");
    assert!(output.contains("PASS: DomainUpDown.Text default is empty"), "DomainUpDown Text default");
    assert!(output.contains("PASS: DomainUpDown.ReadOnly default is False"), "DomainUpDown ReadOnly default");
    assert!(output.contains("PASS: DomainUpDown.Wrap default is False"), "DomainUpDown Wrap default");
    assert!(output.contains("PASS: DomainUpDown.Sorted default is False"), "DomainUpDown Sorted default");
    assert!(output.contains("PASS: DomainUpDown.Items.Count after Add is 3"), "DomainUpDown Items.Count");
    assert!(output.contains("PASS: DomainUpDown.DownButton sets SelectedIndex to 0"), "DomainUpDown DownButton first");
    assert!(output.contains("PASS: DomainUpDown.Text is Red after first DownButton"), "DomainUpDown Text Red");
    assert!(output.contains("PASS: DomainUpDown.DownButton advances to index 1"), "DomainUpDown DownButton index 1");
    assert!(output.contains("PASS: DomainUpDown.Text is Green"), "DomainUpDown Text Green");
    assert!(output.contains("PASS: DomainUpDown.DownButton advances to index 2"), "DomainUpDown DownButton index 2");
    assert!(output.contains("PASS: DomainUpDown.Text is Blue"), "DomainUpDown Text Blue");
    assert!(output.contains("PASS: DomainUpDown.DownButton clamps at last item"), "DomainUpDown DownButton clamp");
    assert!(output.contains("PASS: DomainUpDown.UpButton moves to index 1"), "DomainUpDown UpButton index 1");
    assert!(output.contains("PASS: DomainUpDown.Text is Green after UpButton"), "DomainUpDown Text Green UpButton");
    assert!(output.contains("PASS: DomainUpDown.UpButton moves to index 0"), "DomainUpDown UpButton index 0");
    assert!(output.contains("PASS: DomainUpDown.UpButton clamps at first item"), "DomainUpDown UpButton clamp");
    assert!(output.contains("PASS: DomainUpDown.ReadOnly set to True"), "DomainUpDown ReadOnly set");
    assert!(output.contains("PASS: DomainUpDown.Wrap set to True"), "DomainUpDown Wrap set");
    assert!(output.contains("PASS: DomainUpDown.Sorted set to True"), "DomainUpDown Sorted set");

    // ===== BackgroundWorker =====
    assert!(output.contains("PASS: BackgroundWorker.IsBusy default is False"), "BGW IsBusy");
    assert!(output.contains("PASS: BackgroundWorker.CancellationPending default is False"), "BGW CancellationPending");
    assert!(output.contains("PASS: BackgroundWorker.WorkerReportsProgress default is False"), "BGW WorkerReportsProgress default");
    assert!(output.contains("PASS: BackgroundWorker.WorkerSupportsCancellation default is False"), "BGW WorkerSupportsCancellation default");
    assert!(output.contains("PASS: BackgroundWorker.WorkerReportsProgress set to True"), "BGW WorkerReportsProgress set");
    assert!(output.contains("PASS: BackgroundWorker.WorkerSupportsCancellation set to True"), "BGW WorkerSupportsCancellation set");
    assert!(output.contains("PASS: BackgroundWorker.RunWorkerAsync does not crash"), "BGW RunWorkerAsync");
    assert!(output.contains("PASS: BackgroundWorker.CancelAsync does not crash"), "BGW CancelAsync");

    // ===== HelpProvider =====
    assert!(output.contains("PASS: HelpProvider.HelpNamespace default is empty"), "HelpProvider HelpNamespace default");
    assert!(output.contains("PASS: HelpProvider.HelpNamespace set"), "HelpProvider HelpNamespace set");
    assert!(output.contains("PASS: HelpProvider.SetHelpString does not crash"), "HelpProvider SetHelpString");

    // ===== PrintDialog =====
    assert!(output.contains("PASS: PrintDialog.AllowPrintToFile default is True"), "PrintDialog AllowPrintToFile");
    assert!(output.contains("PASS: PrintDialog.AllowSelection default is False"), "PrintDialog AllowSelection");
    assert!(output.contains("PASS: PrintDialog.AllowSomePages default is False"), "PrintDialog AllowSomePages default");
    assert!(output.contains("PASS: PrintDialog.PrintToFile default is False"), "PrintDialog PrintToFile");
    assert!(output.contains("PASS: PrintDialog.ShowHelp default is False"), "PrintDialog ShowHelp");
    assert!(output.contains("PASS: PrintDialog.ShowNetwork default is True"), "PrintDialog ShowNetwork");
    assert!(output.contains("PASS: PrintDialog.AllowSomePages set to True"), "PrintDialog AllowSomePages set");
    assert!(output.contains("PASS: PrintDialog.ShowDialog returns DialogResult.Cancel in headless mode"), "PrintDialog ShowDialog");

    // ===== PrintPreviewDialog =====
    assert!(output.contains("PASS: PrintPreviewDialog created successfully"), "PrintPreviewDialog created");
    assert!(output.contains("PASS: PrintPreviewDialog.ShowDialog returns DialogResult.Cancel in headless mode"), "PrintPreviewDialog ShowDialog");

    // ===== PageSetupDialog =====
    assert!(output.contains("PASS: PageSetupDialog.AllowMargins default is True"), "PageSetupDialog AllowMargins");
    assert!(output.contains("PASS: PageSetupDialog.AllowOrientation default is True"), "PageSetupDialog AllowOrientation");
    assert!(output.contains("PASS: PageSetupDialog.AllowPaper default is True"), "PageSetupDialog AllowPaper");
    assert!(output.contains("PASS: PageSetupDialog.AllowPrinter default is False"), "PageSetupDialog AllowPrinter default");
    assert!(output.contains("PASS: PageSetupDialog.AllowPrinter set to True"), "PageSetupDialog AllowPrinter set");
    assert!(output.contains("PASS: PageSetupDialog.ShowDialog returns DialogResult.Cancel in headless mode"), "PageSetupDialog ShowDialog");

    // ===== PropertyGrid =====
    assert!(output.contains("PASS: PropertyGrid.Text default is empty"), "PropertyGrid Text");
    assert!(output.contains("PASS: PropertyGrid.Visible default is True"), "PropertyGrid Visible");
    assert!(output.contains("PASS: PropertyGrid.Enabled default is True"), "PropertyGrid Enabled");
    assert!(output.contains("PASS: PropertyGrid.SelectedObject setter does not crash"), "PropertyGrid SelectedObject");

    // ===== Splitter =====
    assert!(output.contains("PASS: Splitter.Visible default is True"), "Splitter Visible");
    assert!(output.contains("PASS: Splitter.Enabled default is True"), "Splitter Enabled");
    assert!(output.contains("PASS: Splitter.Dock set to Right"), "Splitter Dock set");

    // ===== DataGrid (legacy) =====
    assert!(output.contains("PASS: DataGrid.Visible default is True"), "DataGrid Visible");
    assert!(output.contains("PASS: DataGrid.Enabled default is True"), "DataGrid Enabled");
    assert!(output.contains("PASS: DataGrid.Text default is empty"), "DataGrid Text");

    // ===== UserControl =====
    assert!(output.contains("PASS: UserControl.Visible default is True"), "UserControl Visible");
    assert!(output.contains("PASS: UserControl.Enabled default is True"), "UserControl Enabled");

    // ===== ToolStrip sub-items =====
    assert!(output.contains("PASS: ToolStripButton.Text default is empty"), "ToolStripButton Text default");
    assert!(output.contains("PASS: ToolStripButton.Enabled default is True"), "ToolStripButton Enabled");
    assert!(output.contains("PASS: ToolStripButton.Visible default is True"), "ToolStripButton Visible");
    assert!(output.contains("PASS: ToolStripButton.Text set to Save"), "ToolStripButton Text set");
    assert!(output.contains("PASS: ToolStripButton.ToolTipText set"), "ToolStripButton ToolTipText");
    assert!(output.contains("PASS: ToolStripLabel.Text default is empty"), "ToolStripLabel Text default");
    assert!(output.contains("PASS: ToolStripLabel.Text set to Status:"), "ToolStripLabel Text set");
    assert!(output.contains("PASS: ToolStripSeparator.Visible default is True"), "ToolStripSeparator Visible");
    assert!(output.contains("PASS: ToolStripComboBox.Text default is empty"), "ToolStripComboBox Text default");
    assert!(output.contains("PASS: ToolStripComboBox.Items not Nothing"), "ToolStripComboBox Items");
    assert!(output.contains("PASS: ToolStripComboBox.Items.Count after Add is 2"), "ToolStripComboBox Items.Count");
    assert!(output.contains("PASS: ToolStripTextBox.Text default is empty"), "ToolStripTextBox Text default");
    assert!(output.contains("PASS: ToolStripTextBox.Text set"), "ToolStripTextBox Text set");
    assert!(output.contains("PASS: ToolStripProgressBar.Value default is 0"), "ToolStripProgressBar Value default");
    assert!(output.contains("PASS: ToolStripProgressBar.Minimum default is 0"), "ToolStripProgressBar Minimum");
    assert!(output.contains("PASS: ToolStripProgressBar.Maximum default is 100"), "ToolStripProgressBar Maximum");
    assert!(output.contains("PASS: ToolStripProgressBar.Value set to 50"), "ToolStripProgressBar Value set");
    assert!(output.contains("PASS: ToolStripDropDownButton.Text default is empty"), "ToolStripDropDownButton Text default");
    assert!(output.contains("PASS: ToolStripDropDownButton.Text set to File"), "ToolStripDropDownButton Text set");
    assert!(output.contains("PASS: ToolStripSplitButton.Text default is empty"), "ToolStripSplitButton Text default");
    assert!(output.contains("PASS: ToolStripSplitButton.Text set to New"), "ToolStripSplitButton Text set");

    // ===== SqlConnection / OleDbConnection =====
    assert!(output.contains("PASS: SqlConnection.ConnectionString set"), "SqlConnection ConnectionString");
    assert!(output.contains("PASS: SqlConnection created successfully"), "SqlConnection created");
    assert!(output.contains("PASS: OleDbConnection.ConnectionString set"), "OleDbConnection ConnectionString");
    assert!(output.contains("PASS: OleDbConnection created successfully"), "OleDbConnection created");

    // ===== PrintPreviewControl =====
    assert!(output.contains("PASS: PrintPreviewControl.Visible default is True"), "PrintPreviewControl Visible");
    assert!(output.contains("PASS: PrintPreviewControl.Enabled default is True"), "PrintPreviewControl Enabled");
    assert!(output.contains("PASS: PrintPreviewControl created successfully"), "PrintPreviewControl created");

    assert_eq!(fail_count, 0, "Expected 0 failures but got {}\nFull output:\n{}", fail_count, output);
    assert!(pass_count >= 70, "Expected at least 70 passes but got {}", pass_count);
}
