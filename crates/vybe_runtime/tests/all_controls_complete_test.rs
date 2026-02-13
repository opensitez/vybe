use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_all_controls_complete() {
    let source = std::fs::read_to_string("../../tests/test_all_controls_complete.vb")
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

    // Count passes and failures
    let pass_count = output.matches("PASS:").count();
    let fail_count = output.matches("FAIL:").count();
    println!("Passes: {}, Failures: {}", pass_count, fail_count);

    // ===== BUTTON =====
    assert!(output.contains("PASS: Button.FlatStyle"), "Button.FlatStyle");
    assert!(output.contains("PASS: Button.DialogResult"), "Button.DialogResult default");
    assert!(output.contains("PASS: Button.TextAlign"), "Button.TextAlign");
    assert!(output.contains("PASS: Button.AutoEllipsis"), "Button.AutoEllipsis");
    assert!(output.contains("PASS: Button.UseMnemonic"), "Button.UseMnemonic");
    assert!(output.contains("PASS: Button.Text set"), "Button.Text set");
    assert!(output.contains("PASS: Button.DialogResult set"), "Button.DialogResult set");
    assert!(output.contains("PASS: Button.Enabled set"), "Button.Enabled set");

    // ===== LABEL =====
    assert!(output.contains("PASS: Label.AutoSize"), "Label.AutoSize");
    assert!(output.contains("PASS: Label.TextAlign"), "Label.TextAlign");
    assert!(output.contains("PASS: Label.BorderStyle"), "Label.BorderStyle");
    assert!(output.contains("PASS: Label.AutoEllipsis"), "Label.AutoEllipsis");
    assert!(output.contains("PASS: Label.UseMnemonic"), "Label.UseMnemonic");
    assert!(output.contains("PASS: Label.Text set"), "Label.Text set");

    // ===== CHECKBOX =====
    assert!(output.contains("PASS: CheckBox.Checked"), "CheckBox.Checked default");
    assert!(output.contains("PASS: CheckBox.CheckState"), "CheckBox.CheckState");
    assert!(output.contains("PASS: CheckBox.ThreeState"), "CheckBox.ThreeState default");
    assert!(output.contains("PASS: CheckBox.AutoCheck"), "CheckBox.AutoCheck");
    assert!(output.contains("PASS: CheckBox.CheckAlign"), "CheckBox.CheckAlign");
    assert!(output.contains("PASS: CheckBox.Appearance"), "CheckBox.Appearance");
    assert!(output.contains("PASS: CheckBox.Checked set"), "CheckBox.Checked set");
    assert!(output.contains("PASS: CheckBox.ThreeState set"), "CheckBox.ThreeState set");

    // ===== RADIOBUTTON =====
    assert!(output.contains("PASS: RadioButton.Checked"), "RadioButton.Checked default");
    assert!(output.contains("PASS: RadioButton.AutoCheck"), "RadioButton.AutoCheck");
    assert!(output.contains("PASS: RadioButton.CheckAlign"), "RadioButton.CheckAlign");
    assert!(output.contains("PASS: RadioButton.Appearance"), "RadioButton.Appearance");
    assert!(output.contains("PASS: RadioButton.Checked set"), "RadioButton.Checked set");

    // ===== GROUPBOX =====
    assert!(output.contains("PASS: GroupBox.FlatStyle"), "GroupBox.FlatStyle");
    assert!(output.contains("PASS: GroupBox.AutoSize"), "GroupBox.AutoSize");
    assert!(output.contains("PASS: GroupBox.Padding"), "GroupBox.Padding");
    assert!(output.contains("PASS: GroupBox.Text set"), "GroupBox.Text set");

    // ===== PANEL =====
    assert!(output.contains("PASS: Panel.BorderStyle"), "Panel.BorderStyle default");
    assert!(output.contains("PASS: Panel.AutoSize"), "Panel.AutoSize");
    assert!(output.contains("PASS: Panel.AutoScroll"), "Panel.AutoScroll");
    assert!(output.contains("PASS: Panel.Padding"), "Panel.Padding");
    assert!(output.contains("PASS: Panel.BorderStyle set"), "Panel.BorderStyle set");

    // ===== PICTUREBOX =====
    assert!(output.contains("PASS: PictureBox.SizeMode"), "PictureBox.SizeMode default");
    assert!(output.contains("PASS: PictureBox.BorderStyle"), "PictureBox.BorderStyle");
    assert!(output.contains("PASS: PictureBox.WaitOnLoad"), "PictureBox.WaitOnLoad");
    assert!(output.contains("PASS: PictureBox.SizeMode set"), "PictureBox.SizeMode set");

    // ===== TEXTBOX =====
    assert!(output.contains("PASS: TextBox.ReadOnly"), "TextBox.ReadOnly default");
    assert!(output.contains("PASS: TextBox.Multiline"), "TextBox.Multiline default");
    assert!(output.contains("PASS: TextBox.MaxLength"), "TextBox.MaxLength");
    assert!(output.contains("PASS: TextBox.WordWrap"), "TextBox.WordWrap");
    assert!(output.contains("PASS: TextBox.AcceptsReturn"), "TextBox.AcceptsReturn");
    assert!(output.contains("PASS: TextBox.AcceptsTab"), "TextBox.AcceptsTab");
    assert!(output.contains("PASS: TextBox.CharacterCasing"), "TextBox.CharacterCasing");
    assert!(output.contains("PASS: TextBox.HideSelection"), "TextBox.HideSelection");
    assert!(output.contains("PASS: TextBox.BorderStyle"), "TextBox.BorderStyle");
    assert!(output.contains("PASS: TextBox.Text set"), "TextBox.Text set");

    // ===== COMBOBOX =====
    assert!(output.contains("PASS: ComboBox.SelectedIndex"), "ComboBox.SelectedIndex");
    assert!(output.contains("PASS: ComboBox.DropDownStyle"), "ComboBox.DropDownStyle");
    assert!(output.contains("PASS: ComboBox.MaxDropDownItems"), "ComboBox.MaxDropDownItems");
    assert!(output.contains("PASS: ComboBox.Sorted"), "ComboBox.Sorted");
    assert!(output.contains("PASS: ComboBox.Items.Count"), "ComboBox.Items.Count");

    // ===== LISTBOX =====
    assert!(output.contains("PASS: ListBox.SelectedIndex"), "ListBox.SelectedIndex");
    assert!(output.contains("PASS: ListBox.SelectionMode"), "ListBox.SelectionMode");
    assert!(output.contains("PASS: ListBox.Sorted"), "ListBox.Sorted");
    assert!(output.contains("PASS: ListBox.Items.Count"), "ListBox.Items.Count");

    // ===== RICHTEXTBOX =====
    assert!(output.contains("PASS: RichTextBox.Multiline"), "RichTextBox.Multiline");
    assert!(output.contains("PASS: RichTextBox.WordWrap"), "RichTextBox.WordWrap");
    assert!(output.contains("PASS: RichTextBox.DetectUrls"), "RichTextBox.DetectUrls");
    assert!(output.contains("PASS: RichTextBox.Text set"), "RichTextBox.Text set");

    // ===== PROGRESSBAR =====
    assert!(output.contains("PASS: ProgressBar.Value"), "ProgressBar.Value default");
    assert!(output.contains("PASS: ProgressBar.Maximum"), "ProgressBar.Maximum default");
    assert!(output.contains("PASS: ProgressBar.Step"), "ProgressBar.Step");
    assert!(output.contains("PASS: ProgressBar.Value set"), "ProgressBar.Value set");

    // ===== NUMERICUPDOWN =====
    assert!(output.contains("PASS: NumericUpDown.Value"), "NumericUpDown.Value default");
    assert!(output.contains("PASS: NumericUpDown.Maximum"), "NumericUpDown.Maximum");
    assert!(output.contains("PASS: NumericUpDown.Increment"), "NumericUpDown.Increment");
    assert!(output.contains("PASS: NumericUpDown.DecimalPlaces"), "NumericUpDown.DecimalPlaces");
    assert!(output.contains("PASS: NumericUpDown.Value set"), "NumericUpDown.Value set");

    // ===== TREEVIEW =====
    assert!(output.contains("PASS: TreeView.CheckBoxes"), "TreeView.CheckBoxes");
    assert!(output.contains("PASS: TreeView.ShowLines"), "TreeView.ShowLines");
    assert!(output.contains("PASS: TreeView.ShowRootLines"), "TreeView.ShowRootLines");
    assert!(output.contains("PASS: TreeView.ShowPlusMinus"), "TreeView.ShowPlusMinus");
    assert!(output.contains("PASS: TreeView.Sorted"), "TreeView.Sorted");

    // ===== LISTVIEW =====
    assert!(output.contains("PASS: ListView.View"), "ListView.View");
    assert!(output.contains("PASS: ListView.FullRowSelect"), "ListView.FullRowSelect");
    assert!(output.contains("PASS: ListView.GridLines"), "ListView.GridLines");
    assert!(output.contains("PASS: ListView.MultiSelect"), "ListView.MultiSelect");

    // ===== DATAGRIDVIEW =====
    assert!(output.contains("PASS: DataGridView.AllowUserToAddRows"), "DGV.AllowUserToAddRows");
    assert!(output.contains("PASS: DataGridView.ReadOnly"), "DGV.ReadOnly");
    assert!(output.contains("PASS: DataGridView.MultiSelect"), "DGV.MultiSelect");
    assert!(output.contains("PASS: DataGridView.SelectionMode"), "DGV.SelectionMode");
    assert!(output.contains("PASS: DataGridView.EditMode"), "DGV.EditMode");

    // ===== TABCONTROL =====
    assert!(output.contains("PASS: TabControl.Alignment"), "TabControl.Alignment");
    assert!(output.contains("PASS: TabControl.Appearance"), "TabControl.Appearance");
    assert!(output.contains("PASS: TabControl.Multiline"), "TabControl.Multiline");

    // ===== TABPAGE =====
    assert!(output.contains("PASS: TabPage.UseVisualStyleBackColor"), "TabPage.UseVisualStyleBackColor");
    assert!(output.contains("PASS: TabPage.Padding"), "TabPage.Padding");
    assert!(output.contains("PASS: TabPage.Text set"), "TabPage.Text set");

    // ===== MENUSTRIP =====
    assert!(output.contains("PASS: MenuStrip.Dock"), "MenuStrip.Dock");
    assert!(output.contains("PASS: MenuStrip.Stretch"), "MenuStrip.Stretch");

    // ===== STATUSSTRIP =====
    assert!(output.contains("PASS: StatusStrip.Dock"), "StatusStrip.Dock");
    assert!(output.contains("PASS: StatusStrip.SizingGrip"), "StatusStrip.SizingGrip");
    assert!(output.contains("PASS: StatusStrip.Stretch"), "StatusStrip.Stretch");

    // ===== TOOLSTRIPSTATUSLABEL =====
    assert!(output.contains("PASS: ToolStripStatusLabel.Spring"), "ToolStripStatusLabel.Spring");
    assert!(output.contains("PASS: ToolStripStatusLabel.AutoSize"), "ToolStripStatusLabel.AutoSize");
    assert!(output.contains("PASS: ToolStripStatusLabel.Text set"), "ToolStripStatusLabel.Text set");

    // ===== TOOLSTRIPMENUITEM =====
    assert!(output.contains("PASS: ToolStripMenuItem.Checked"), "ToolStripMenuItem.Checked default");
    assert!(output.contains("PASS: ToolStripMenuItem.CheckOnClick"), "ToolStripMenuItem.CheckOnClick");
    assert!(output.contains("PASS: ToolStripMenuItem.Checked set"), "ToolStripMenuItem.Checked set");

    // ===== DATETIMEPICKER =====
    assert!(output.contains("PASS: DateTimePicker.Format"), "DateTimePicker.Format default");
    assert!(output.contains("PASS: DateTimePicker.ShowCheckBox"), "DateTimePicker.ShowCheckBox");
    assert!(output.contains("PASS: DateTimePicker.ShowUpDown"), "DateTimePicker.ShowUpDown");
    assert!(output.contains("PASS: DateTimePicker.Format set"), "DateTimePicker.Format set");

    // ===== LINKLABEL =====
    assert!(output.contains("PASS: LinkLabel.LinkColor"), "LinkLabel.LinkColor");
    assert!(output.contains("PASS: LinkLabel.LinkVisited"), "LinkLabel.LinkVisited default");
    assert!(output.contains("PASS: LinkLabel.LinkVisited set"), "LinkLabel.LinkVisited set");

    // ===== TOOLSTRIP =====
    assert!(output.contains("PASS: ToolStrip.Dock"), "ToolStrip.Dock");
    assert!(output.contains("PASS: ToolStrip.ShowItemToolTips"), "ToolStrip.ShowItemToolTips");
    assert!(output.contains("PASS: ToolStrip.LayoutStyle"), "ToolStrip.LayoutStyle");

    // ===== TRACKBAR =====
    assert!(output.contains("PASS: TrackBar.Value"), "TrackBar.Value default");
    assert!(output.contains("PASS: TrackBar.Maximum"), "TrackBar.Maximum default");
    assert!(output.contains("PASS: TrackBar.TickFrequency"), "TrackBar.TickFrequency");
    assert!(output.contains("PASS: TrackBar.Value set"), "TrackBar.Value set");

    // ===== MASKEDTEXTBOX =====
    assert!(output.contains("PASS: MaskedTextBox.PromptChar"), "MaskedTextBox.PromptChar");
    assert!(output.contains("PASS: MaskedTextBox.SkipLiterals"), "MaskedTextBox.SkipLiterals");
    assert!(output.contains("PASS: MaskedTextBox.Mask set"), "MaskedTextBox.Mask set");

    // ===== SPLITCONTAINER =====
    assert!(output.contains("PASS: SplitContainer.Orientation"), "SplitContainer.Orientation");
    assert!(output.contains("PASS: SplitContainer.SplitterDistance"), "SplitContainer.SplitterDistance default");
    assert!(output.contains("PASS: SplitContainer.IsSplitterFixed"), "SplitContainer.IsSplitterFixed");
    assert!(output.contains("PASS: SplitContainer.SplitterDistance set"), "SplitContainer.SplitterDistance set");

    // ===== FLOWLAYOUTPANEL =====
    assert!(output.contains("PASS: FlowLayoutPanel.FlowDirection"), "FlowLayoutPanel.FlowDirection default");
    assert!(output.contains("PASS: FlowLayoutPanel.WrapContents"), "FlowLayoutPanel.WrapContents");
    assert!(output.contains("PASS: FlowLayoutPanel.FlowDirection set"), "FlowLayoutPanel.FlowDirection set");

    // ===== TABLELAYOUTPANEL =====
    assert!(output.contains("PASS: TableLayoutPanel.ColumnCount"), "TableLayoutPanel.ColumnCount default");
    assert!(output.contains("PASS: TableLayoutPanel.RowCount"), "TableLayoutPanel.RowCount");
    assert!(output.contains("PASS: TableLayoutPanel.GrowStyle"), "TableLayoutPanel.GrowStyle");
    assert!(output.contains("PASS: TableLayoutPanel.ColumnCount set"), "TableLayoutPanel.ColumnCount set");

    // ===== MONTHCALENDAR =====
    assert!(output.contains("PASS: MonthCalendar.ShowToday"), "MonthCalendar.ShowToday");
    assert!(output.contains("PASS: MonthCalendar.MaxSelectionCount"), "MonthCalendar.MaxSelectionCount default");
    assert!(output.contains("PASS: MonthCalendar.ShowWeekNumbers set"), "MonthCalendar.ShowWeekNumbers set");

    // ===== SCROLLBARS =====
    assert!(output.contains("PASS: HScrollBar.Value"), "HScrollBar.Value default");
    assert!(output.contains("PASS: HScrollBar.LargeChange"), "HScrollBar.LargeChange");
    assert!(output.contains("PASS: VScrollBar.Value"), "VScrollBar.Value default");
    assert!(output.contains("PASS: VScrollBar.LargeChange"), "VScrollBar.LargeChange");

    // ===== TOOLTIP =====
    assert!(output.contains("PASS: ToolTip.Active"), "ToolTip.Active");
    assert!(output.contains("PASS: ToolTip.AutoPopDelay"), "ToolTip.AutoPopDelay");
    assert!(output.contains("PASS: ToolTip.SetToolTip/GetToolTip"), "ToolTip.SetToolTip/GetToolTip");
    assert!(output.contains("PASS: ToolTip.RemoveAll"), "ToolTip.RemoveAll");

    // ===== WEBBROWSER =====
    assert!(output.contains("PASS: WebBrowser.CanGoBack"), "WebBrowser.CanGoBack");
    assert!(output.contains("PASS: WebBrowser.ReadyState"), "WebBrowser.ReadyState");
    assert!(output.contains("PASS: WebBrowser.AllowNavigation"), "WebBrowser.AllowNavigation");

    // Final: no failures
    assert_eq!(fail_count, 0, "Expected zero FAIL results, got {}", fail_count);
    assert!(pass_count >= 150, "Expected at least 150 PASS results, got {}", pass_count);
}
