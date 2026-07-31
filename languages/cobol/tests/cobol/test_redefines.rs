use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_redefines_basic() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "AAAA".
01 WS-ALIAS REDEFINES WS-BASE PIC X(4).
"#,
        r#"
    MOVE "BBBB" TO WS-ALIAS.
    DISPLAY WS-BASE.
"#,
    ));
    assert_eq!(output, vec!["BBBB"]);
}

#[test]
fn test_redefines_types() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "1234".
01 WS-ALIAS REDEFINES WS-BASE PIC 9(4).
"#,
        r#"
    DISPLAY WS-ALIAS.
"#,
    ));
    assert_eq!(output, vec!["1234"]);
}

#[test]
fn test_redefines_partial() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "ABCD".
01 WS-ALIAS REDEFINES WS-BASE PIC X(2).
"#,
        r#"
    DISPLAY WS-ALIAS.
"#,
    ));
    assert_eq!(output, vec!["AB"]);
}

#[test]
fn test_redefines_group() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(6) VALUE "123456".
01 WS-ALIAS REDEFINES WS-BASE.
   05 WS-A PIC X(3).
   05 WS-B PIC X(3).
"#,
        r#"
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["123", "456"]);
}

#[test]
fn test_redefines_multiple() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "AAAA".
01 WS-ALIAS1 REDEFINES WS-BASE PIC 9(4).
01 WS-ALIAS2 REDEFINES WS-BASE PIC X(4).
"#,
        r#"
    MOVE "4444" TO WS-ALIAS2.
    DISPLAY WS-ALIAS1.
    DISPLAY WS-ALIAS2.
"#,
    ));
    assert_eq!(output, vec!["4444", "4444"]);
}

#[test]
fn test_redefines_nested_group_with_numeric() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(8) VALUE "12345678".
01 WS-REDEF REDEFINES WS-BASE.
   05 WS-N1 PIC 9(4).
   05 WS-N2 PIC 9(4).
"#,
        r#"
    ADD 1 TO WS-N1.
    ADD 1 TO WS-N2.
    DISPLAY WS-N1.
    DISPLAY WS-N2.
"#,
    ));
    assert_eq!(output, vec!["1235", "5679"]);
}

#[test]
fn test_redefines_alias_roundtrip_to_base() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE SPACES.
01 WS-ALIAS REDEFINES WS-BASE PIC 9(4).
"#,
        r#"
    MOVE 77 TO WS-ALIAS.
    DISPLAY WS-BASE.
"#,
    ));
    assert_eq!(output, vec!["0077"]);
}

#[test]
fn test_redefines_partial_view_consistency() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(6) VALUE "ABCDEF".
01 WS-ALIAS REDEFINES WS-BASE PIC X(2).
01 WS-TAIL PIC X(4) VALUE "WXYZ".
"#,
        r#"
    MOVE WS-TAIL TO WS-BASE(3:4).
    DISPLAY WS-ALIAS.
    DISPLAY WS-BASE(3:4).
"#,
    ));
    assert_eq!(output, vec!["AB", "WXYZ"]);
}
