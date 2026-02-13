use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_winforms_features() {
    let source = std::fs::read_to_string("../../tests/test_winforms_features.vb")
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

    // TextBox properties
    assert!(output.contains("TextBox.ReadOnly: True"), "TextBox.ReadOnly");
    assert!(output.contains("TextBox.Multiline: True"), "TextBox.Multiline");
    assert!(output.contains("TextBox.PasswordChar: *"), "TextBox.PasswordChar");

    // ComboBox Items
    assert!(output.contains("ComboBox.Items.Count: 3"), "ComboBox Items count");
    assert!(output.contains("ComboBox.SelectedIndex: 1"), "ComboBox SelectedIndex");

    // ListBox Items
    assert!(output.contains("ListBox.Items.Count: 3"), "ListBox Items count");
    assert!(output.contains("ListBox after Clear: 0"), "ListBox Clear");

    // ProgressBar
    assert!(output.contains("ProgressBar.Value: 10"), "ProgressBar Value");
    assert!(output.contains("ProgressBar.Maximum: 100"), "ProgressBar Maximum");
    assert!(output.contains("ProgressBar after PerformStep: 20"), "ProgressBar PerformStep");

    // NumericUpDown
    assert!(output.contains("NumericUpDown.Value: 10"), "NumericUpDown Value");
    assert!(output.contains("NumericUpDown after UpButton: 15"), "NumericUpDown UpButton");
    assert!(output.contains("NumericUpDown after DownButton: 10"), "NumericUpDown DownButton");

    // TreeView Nodes
    assert!(output.contains("TreeView.Nodes.Count: 2"), "TreeView Nodes count");

    // ListView Items/Columns
    assert!(output.contains("ListView.Columns.Count: 2"), "ListView Columns count");
    assert!(output.contains("ListView.Items.Count: 2"), "ListView Items count");

    // DataGridView Rows/Columns
    assert!(output.contains("DataGridView.ColumnCount: 2"), "DataGridView ColumnCount");
    assert!(output.contains("DataGridView.RowCount: 1"), "DataGridView RowCount");

    // TabControl TabPages
    assert!(output.contains("TabControl.TabCount: 3"), "TabControl TabCount");
    assert!(output.contains("TabControl.SelectedIndex: 0"), "TabControl initial SelectedIndex");
    assert!(output.contains("TabControl.SelectedIndex after set: 2"), "TabControl set SelectedIndex");

    // MenuStrip Items
    assert!(output.contains("MenuStrip.Items.Count: 2"), "MenuStrip Items count");

    // ToolStripMenuItem DropDownItems
    assert!(output.contains("ToolStripMenuItem.Text: File"), "ToolStripMenuItem Text");
    assert!(output.contains("ToolStripMenuItem.DropDownItems.Count: 3"), "ToolStripMenuItem DropDownItems count");

    // Layout properties
    assert!(output.contains("Panel.Dock: 5"), "Panel Dock Fill");
    assert!(output.contains("Panel.Anchor: 5"), "Panel Anchor Top|Left");

    assert!(output.contains("=== WinForms Feature Tests Complete ==="), "Completed");
}
