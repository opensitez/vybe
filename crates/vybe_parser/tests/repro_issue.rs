use vybe_parser::parse_program;

#[test]
fn test_repro_simple() {
    let code = "Sub btn1_Click()\n    MsgBox \"Hi\"\nEnd Sub";
    let result = parse_program(code);
    assert!(result.is_ok(), "{:?}", result.err());
}

#[test]
fn test_repro_with_newlines() {
    let code = "
Sub btn1_Click()
    MsgBox \"Hi\"
End Sub
";
    let result = parse_program(code);
    assert!(result.is_ok(), "{:?}", result.err());
}

#[test]
fn test_repro_end_sub_space() {
    let code = "Sub btn1_Click()\n    MsgBox \"Hi\"\nEnd  Sub";
    let result = parse_program(code);
    assert!(result.is_ok(), "{:?}", result.err());
}
