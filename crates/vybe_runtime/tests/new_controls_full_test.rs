use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_new_controls_full() {
    let source = std::fs::read_to_string("../../tests/test_new_controls_full.vb")
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

    // MonthCalendar tests
    assert!(output.contains("PASS: MonthCalendar.ShowToday default is True"), "MonthCalendar ShowToday");
    assert!(output.contains("PASS: MonthCalendar.ShowTodayCircle default is True"), "MonthCalendar ShowTodayCircle");
    assert!(output.contains("PASS: MonthCalendar.ShowWeekNumbers default is False"), "MonthCalendar ShowWeekNumbers default");
    assert!(output.contains("PASS: MonthCalendar.MaxSelectionCount default is 7"), "MonthCalendar MaxSelectionCount");
    assert!(output.contains("PASS: MonthCalendar.FirstDayOfWeek default is Default"), "MonthCalendar FirstDayOfWeek");
    assert!(output.contains("PASS: MonthCalendar.ScrollChange default is 1"), "MonthCalendar ScrollChange");
    assert!(output.contains("PASS: MonthCalendar.ShowWeekNumbers set to True"), "MonthCalendar ShowWeekNumbers set");
    assert!(output.contains("PASS: MonthCalendar.MaxSelectionCount set to 14"), "MonthCalendar MaxSelectionCount set");

    // HScrollBar tests
    assert!(output.contains("PASS: HScrollBar.Value default is 0"), "HScrollBar Value default");
    assert!(output.contains("PASS: HScrollBar.Minimum default is 0"), "HScrollBar Minimum default");
    assert!(output.contains("PASS: HScrollBar.Maximum default is 100"), "HScrollBar Maximum default");
    assert!(output.contains("PASS: HScrollBar.SmallChange default is 1"), "HScrollBar SmallChange default");
    assert!(output.contains("PASS: HScrollBar.LargeChange default is 10"), "HScrollBar LargeChange default");
    assert!(output.contains("PASS: HScrollBar.Value set to 50"), "HScrollBar Value set");
    assert!(output.contains("PASS: HScrollBar.Maximum set to 200"), "HScrollBar Maximum set");

    // VScrollBar tests
    assert!(output.contains("PASS: VScrollBar.Value default is 0"), "VScrollBar Value default");
    assert!(output.contains("PASS: VScrollBar.Maximum default is 100"), "VScrollBar Maximum default");
    assert!(output.contains("PASS: VScrollBar.LargeChange default is 10"), "VScrollBar LargeChange default");
    assert!(output.contains("PASS: VScrollBar.Value set to 75"), "VScrollBar Value set");

    // ToolTip tests
    assert!(output.contains("PASS: ToolTip.Active default is True"), "ToolTip Active default");
    assert!(output.contains("PASS: ToolTip.AutoPopDelay default is 5000"), "ToolTip AutoPopDelay default");
    assert!(output.contains("PASS: ToolTip.InitialDelay default is 500"), "ToolTip InitialDelay default");
    assert!(output.contains("PASS: ToolTip.ReshowDelay default is 100"), "ToolTip ReshowDelay default");
    assert!(output.contains("PASS: ToolTip.ShowAlways default is False"), "ToolTip ShowAlways default");
    assert!(output.contains("PASS: ToolTip.UseFading default is True"), "ToolTip UseFading");
    assert!(output.contains("PASS: ToolTip.UseAnimation default is True"), "ToolTip UseAnimation");
    assert!(output.contains("PASS: ToolTip.AutoPopDelay set to 10000"), "ToolTip AutoPopDelay set");
    assert!(output.contains("PASS: ToolTip.SetToolTip/GetToolTip works"), "ToolTip SetToolTip/GetToolTip");
    assert!(output.contains("PASS: ToolTip.RemoveAll clears tooltips"), "ToolTip RemoveAll");

    // Enhanced existing control tests
    assert!(output.contains("PASS: TextBox.AcceptsReturn default is False"), "TextBox AcceptsReturn");
    assert!(output.contains("PASS: TextBox.AcceptsTab default is False"), "TextBox AcceptsTab");
    assert!(output.contains("PASS: TextBox.CharacterCasing default is Normal"), "TextBox CharacterCasing");
    assert!(output.contains("PASS: TextBox.SelectionStart default is 0"), "TextBox SelectionStart");
    assert!(output.contains("PASS: TextBox.HideSelection default is True"), "TextBox HideSelection");

    assert!(output.contains("PASS: RichTextBox.Multiline default is True"), "RichTextBox Multiline");
    assert!(output.contains("PASS: RichTextBox.DetectUrls default is True"), "RichTextBox DetectUrls");
    assert!(output.contains("PASS: RichTextBox.ZoomFactor default is 1.0"), "RichTextBox ZoomFactor");

    assert!(output.contains("PASS: DateTimePicker.ShowUpDown default is False"), "DateTimePicker ShowUpDown");
    assert!(output.contains("PASS: DateTimePicker.DropDownAlign default is Left"), "DateTimePicker DropDownAlign");

    assert!(output.contains("PASS: LinkLabel.ActiveLinkColor default is Red"), "LinkLabel ActiveLinkColor");
    assert!(output.contains("PASS: LinkLabel.LinkBehavior default is SystemDefault"), "LinkLabel LinkBehavior");
    assert!(output.contains("PASS: LinkLabel.AutoSize default is True"), "LinkLabel AutoSize");

    assert!(output.contains("PASS: TrackBar.TickStyle default is BottomRight"), "TrackBar TickStyle");

    assert!(output.contains("PASS: MaskedTextBox.HidePromptOnLeave default is False"), "MaskedTextBox HidePromptOnLeave");
    assert!(output.contains("PASS: MaskedTextBox.AsciiOnly default is False"), "MaskedTextBox AsciiOnly");
    assert!(output.contains("PASS: MaskedTextBox.SkipLiterals default is True"), "MaskedTextBox SkipLiterals");

    assert!(output.contains("PASS: SplitContainer.IsSplitterFixed default is False"), "SplitContainer IsSplitterFixed");
    assert!(output.contains("PASS: SplitContainer.Panel1Collapsed default is False"), "SplitContainer Panel1Collapsed");
    assert!(output.contains("PASS: SplitContainer.Panel1MinSize default is 25"), "SplitContainer Panel1MinSize");
    assert!(output.contains("PASS: SplitContainer.SplitterWidth default is 4"), "SplitContainer SplitterWidth");

    assert!(output.contains("PASS: FlowLayoutPanel.AutoSize default is False"), "FlowLayoutPanel AutoSize");
    assert!(output.contains("PASS: FlowLayoutPanel.BorderStyle default is None"), "FlowLayoutPanel BorderStyle");

    assert!(output.contains("PASS: TableLayoutPanel.CellBorderStyle default is None"), "TableLayoutPanel CellBorderStyle");
    assert!(output.contains("PASS: TableLayoutPanel.GrowStyle default is AddRows"), "TableLayoutPanel GrowStyle");

    assert!(output.contains("PASS: ComboBox.Sorted default is False"), "ComboBox Sorted");
    assert!(output.contains("PASS: ComboBox.MaxDropDownItems default is 8"), "ComboBox MaxDropDownItems");

    assert!(output.contains("PASS: ListBox.Sorted default is False"), "ListBox Sorted");
    assert!(output.contains("PASS: ListBox.IntegralHeight default is True"), "ListBox IntegralHeight");

    assert!(output.contains("PASS: NumericUpDown.Hexadecimal default is False"), "NumericUpDown Hexadecimal");
    assert!(output.contains("PASS: NumericUpDown.ThousandsSeparator default is False"), "NumericUpDown ThousandsSeparator");

    assert!(output.contains("PASS: DataGridView.AutoGenerateColumns default is True"), "DataGridView AutoGenerateColumns");
    assert!(output.contains("PASS: DataGridView.MultiSelect default is True"), "DataGridView MultiSelect");
    assert!(output.contains("PASS: DataGridView.ColumnHeadersVisible default is True"), "DataGridView ColumnHeadersVisible");
    assert!(output.contains("PASS: DataGridView.RowHeadersVisible default is True"), "DataGridView RowHeadersVisible");

    assert!(output.contains("PASS: TabControl.Alignment default is Top"), "TabControl Alignment");
    assert!(output.contains("PASS: TabControl.Appearance default is Normal"), "TabControl Appearance");

    assert!(output.contains("PASS: TreeView.ShowPlusMinus default is True"), "TreeView ShowPlusMinus");
    assert!(output.contains("PASS: TreeView.LabelEdit default is False"), "TreeView LabelEdit");
    assert!(output.contains("PASS: TreeView.Scrollable default is True"), "TreeView Scrollable");

    assert!(output.contains("PASS: ListView.ShowGroups default is False"), "ListView ShowGroups");
    assert!(output.contains("PASS: ListView.Sorting default is None"), "ListView Sorting");
    assert!(output.contains("PASS: ListView.AllowColumnReorder default is False"), "ListView AllowColumnReorder");

    // Ensure no failures
    assert_eq!(fail_count, 0, "Expected 0 failures but got {}", fail_count);
    assert!(pass_count >= 60, "Expected at least 60 passes but got {}", pass_count);
}
