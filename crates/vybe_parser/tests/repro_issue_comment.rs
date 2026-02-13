use vybe_parser::parse_program;

#[test]
fn test_repro_comment_after_end_sub() {
    let code = "Sub btn1_Click()\n    MsgBox \"Hi\"\nEnd Sub ' Comment";
    let result = parse_program(code);
    assert!(result.is_ok(), "{:?}", result.err());
}
