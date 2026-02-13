use vybe_parser::parse_program;

#[test]
fn test_repro_end_su_error() {
    let code = "Sub btn1_Click()\n    MsgBox \"Hi\"\nEnd Su";
    let result = parse_program(code);
    assert!(result.is_err());
    println!("Error: {:?}", result.err());
}
