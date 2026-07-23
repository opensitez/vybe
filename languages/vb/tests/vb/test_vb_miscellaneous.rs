use super::helpers::run_vb;

// Remaining miscellaneous tests to hit exactly 500
#[test]
fn misc_octal_literal() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Dim o = &O10: Console.WriteLine(o): End Sub: End Module"#),
        vec!["8"]
    );
}
#[test]
fn misc_hex_literal() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Dim h = &H10: Console.WriteLine(h): End Sub: End Module"#),
        vec!["16"]
    );
}
#[test]
fn misc_binary_literal() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim b = &B1010: Console.WriteLine(b): End Sub: End Module"#
        ),
        vec!["10"]
    );
}
#[test]
fn misc_date_literal_format() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim d = #8/24/2020 12:30:00 PM#: Console.WriteLine(d.Year): End Sub: End Module"#
        ),
        vec!["2020"]
    );
}
#[test]
fn misc_exponent_literal() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Dim e = 1.5E2: Console.WriteLine(e): End Sub: End Module"#),
        vec!["150"]
    );
}
#[test]
fn misc_underscore_variable() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim _var = 42: Console.WriteLine(_var): End Sub: End Module"#
        ),
        vec!["42"]
    );
}
#[test]
fn misc_bracket_identifier() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim [Class] = 42: Console.WriteLine([Class]): End Sub: End Module"#
        ),
        vec!["42"]
    );
}
#[test]
fn misc_rem_comment() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): REM This is a comment: Console.WriteLine("OK"): End Sub: End Module"#
        ),
        vec!["OK"]
    );
}
#[test]
fn misc_xml_axis_value() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim xml = <R><C>1</C></R>: Console.WriteLine(xml.<C>.Value): End Sub: End Module"#
        ),
        vec!["1"]
    );
}
#[test]
fn misc_mid_statement_advanced() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim s = "Hello": Mid(s, 1, 1) = "C": Console.WriteLine(s): End Sub: End Module"#
        ),
        vec!["Cello"]
    );
}
