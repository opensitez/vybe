use irys_runtime::{Interpreter, RuntimeSideEffect};
use irys_parser::ast::Identifier;
use irys_parser::parse_program;

#[test]
fn test_new_controls() {
    let source = std::fs::read_to_string("../../tests/test_new_controls.vb")
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

    // DateTimePicker
    assert!(output.contains("DateTimePicker created: True"), "DateTimePicker creation");
    assert!(output.contains("Default Format: Long"), "DateTimePicker default Format");
    assert!(output.contains("Default Checked: True"), "DateTimePicker default Checked");
    assert!(output.contains("Default ShowCheckBox: False"), "DateTimePicker default ShowCheckBox");
    assert!(output.contains("Format after set: Short"), "DateTimePicker Format setter");
    assert!(output.contains("CustomFormat: yyyy-MM-dd"), "DateTimePicker CustomFormat");
    assert!(output.contains("Value: 2025-01-15"), "DateTimePicker Value");
    assert!(output.contains("ShowCheckBox: True"), "DateTimePicker ShowCheckBox setter");
    assert!(output.contains("Checked: False"), "DateTimePicker Checked setter");
    assert!(output.contains("DateTimePicker PASS"), "DateTimePicker PASS");

    // LinkLabel
    assert!(output.contains("LinkLabel created: True"), "LinkLabel creation");
    assert!(output.contains("Default LinkColor: #0066cc"), "LinkLabel default LinkColor");
    assert!(output.contains("Default VisitedLinkColor: #800080"), "LinkLabel default VisitedLinkColor");
    assert!(output.contains("Default LinkVisited: False"), "LinkLabel default LinkVisited");
    assert!(output.contains("Text: Click Here"), "LinkLabel Text");
    assert!(output.contains("LinkColor: #FF0000"), "LinkLabel LinkColor setter");
    assert!(output.contains("VisitedLinkColor: #00FF00"), "LinkLabel VisitedLinkColor setter");
    assert!(output.contains("LinkVisited: True"), "LinkLabel LinkVisited setter");
    assert!(output.contains("LinkLabel PASS"), "LinkLabel PASS");

    // ToolStrip
    assert!(output.contains("ToolStrip created: True"), "ToolStrip creation");
    assert!(output.contains("Items exists: True"), "ToolStrip Items collection");
    assert!(output.contains("ToolStrip PASS"), "ToolStrip PASS");

    // TrackBar
    assert!(output.contains("TrackBar created: True"), "TrackBar creation");
    assert!(output.contains("Default Value: 0"), "TrackBar default Value");
    assert!(output.contains("Default Minimum: 0"), "TrackBar default Minimum");
    assert!(output.contains("Default Maximum: 10"), "TrackBar default Maximum");
    assert!(output.contains("Default TickFrequency: 1"), "TrackBar default TickFrequency");
    assert!(output.contains("Default SmallChange: 1"), "TrackBar default SmallChange");
    assert!(output.contains("Default LargeChange: 5"), "TrackBar default LargeChange");
    assert!(output.contains("Default Orientation: Horizontal"), "TrackBar default Orientation");
    assert!(output.contains("Value after set: 5"), "TrackBar Value setter");
    assert!(output.contains("Minimum: 1"), "TrackBar Minimum setter");
    assert!(output.contains("Maximum: 20"), "TrackBar Maximum setter");
    assert!(output.contains("TickFrequency: 2"), "TrackBar TickFrequency setter");
    assert!(output.contains("SmallChange: 2"), "TrackBar SmallChange setter");
    assert!(output.contains("LargeChange: 10"), "TrackBar LargeChange setter");
    assert!(output.contains("Orientation: Vertical"), "TrackBar Orientation setter");
    assert!(output.contains("TrackBar PASS"), "TrackBar PASS");

    // MaskedTextBox
    assert!(output.contains("MaskedTextBox created: True"), "MaskedTextBox creation");
    assert!(output.contains("Default Mask: []"), "MaskedTextBox default Mask");
    assert!(output.contains("Default PromptChar: _"), "MaskedTextBox default PromptChar");
    assert!(output.contains("Text: 555-12-3456"), "MaskedTextBox Text");
    assert!(output.contains("Mask: 000-00-0000"), "MaskedTextBox Mask setter");
    assert!(output.contains("PromptChar: #"), "MaskedTextBox PromptChar setter");
    assert!(output.contains("MaskedTextBox PASS"), "MaskedTextBox PASS");

    // SplitContainer
    assert!(output.contains("SplitContainer created: True"), "SplitContainer creation");
    assert!(output.contains("Default Orientation: Vertical"), "SplitContainer default Orientation");
    assert!(output.contains("Default SplitterDistance: 100"), "SplitContainer default SplitterDistance");
    assert!(output.contains("Orientation: Horizontal"), "SplitContainer Orientation setter");
    assert!(output.contains("SplitterDistance: 200"), "SplitContainer SplitterDistance setter");
    assert!(output.contains("SplitContainer PASS"), "SplitContainer PASS");

    // FlowLayoutPanel
    assert!(output.contains("FlowLayoutPanel created: True"), "FlowLayoutPanel creation");
    assert!(output.contains("Default FlowDirection: LeftToRight"), "FlowLayoutPanel default FlowDirection");
    assert!(output.contains("Default WrapContents: True"), "FlowLayoutPanel default WrapContents");
    assert!(output.contains("FlowDirection: TopDown"), "FlowLayoutPanel FlowDirection setter");
    assert!(output.contains("WrapContents: False"), "FlowLayoutPanel WrapContents setter");
    assert!(output.contains("FlowLayoutPanel PASS"), "FlowLayoutPanel PASS");

    // TableLayoutPanel
    assert!(output.contains("TableLayoutPanel created: True"), "TableLayoutPanel creation");
    assert!(output.contains("Default ColumnCount: 2"), "TableLayoutPanel default ColumnCount");
    assert!(output.contains("Default RowCount: 2"), "TableLayoutPanel default RowCount");
    assert!(output.contains("ColumnCount: 3"), "TableLayoutPanel ColumnCount setter");
    assert!(output.contains("RowCount: 4"), "TableLayoutPanel RowCount setter");
    assert!(output.contains("TableLayoutPanel PASS"), "TableLayoutPanel PASS");

    // StatusStrip
    assert!(output.contains("StatusStrip created: True"), "StatusStrip creation");
    assert!(output.contains("Items exists: True"), "StatusStrip Items collection");
    assert!(output.contains("StatusStrip PASS"), "StatusStrip PASS");

    // ToolStripStatusLabel
    assert!(output.contains("ToolStripStatusLabel created: True"), "ToolStripStatusLabel creation");
    assert!(output.contains("Text: Ready"), "ToolStripStatusLabel Text");
    assert!(output.contains("Spring: True"), "ToolStripStatusLabel Spring");
    assert!(output.contains("ToolStripStatusLabel PASS"), "ToolStripStatusLabel PASS");

    // Final
    assert!(output.contains("ALL NEW CONTROL TESTS PASSED"), "All tests passed");
}
