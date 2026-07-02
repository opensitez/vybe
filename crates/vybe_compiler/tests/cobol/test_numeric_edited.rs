use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_numeric_edited_zero_suppress() {
    let output = run_prints(&p(
        r#"
01 WS-VAL PIC 9(3) VALUE 42.
01 WS-EDIT PIC Z(3) VALUE ZERO.
"#,
        r#"
    MOVE WS-VAL TO WS-EDIT.
    DISPLAY WS-EDIT.
"#,
    ));
    assert_eq!(output, vec![" 42"]);
}

#[test]
fn test_numeric_edited_asterisk() {
    let output = run_prints(&p(
        r#"
01 WS-VAL PIC 9(3) VALUE 42.
01 WS-EDIT PIC *(3) VALUE ZERO.
"#,
        r#"
    MOVE WS-VAL TO WS-EDIT.
    DISPLAY WS-EDIT.
"#,
    ));
    assert_eq!(output, vec!["*42"]);
}

#[test]
fn test_numeric_edited_currency() {
    let output = run_prints(&p(
        r#"
01 WS-VAL PIC 9(3) VALUE 42.
01 WS-EDIT PIC $9(3) VALUE ZERO.
"#,
        r#"
    MOVE WS-VAL TO WS-EDIT.
    DISPLAY WS-EDIT.
"#,
    ));
    assert_eq!(output, vec!["$042"]);
}

#[test]
fn test_numeric_edited_signs() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC S9(3) VALUE -42.
01 WS-EDIT1 PIC +9(3) VALUE ZERO.
01 WS-EDIT2 PIC -9(3) VALUE ZERO.
"#,
        r#"
    MOVE WS-SRC TO WS-EDIT1.
    DISPLAY WS-EDIT1.
    MOVE WS-SRC TO WS-EDIT2.
    DISPLAY WS-EDIT2.
"#,
    ));
    assert_eq!(output, vec!["-042", "-042"]);
}

#[test]
fn test_numeric_edited_point_comma() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(5)V99 VALUE 12345.67.
01 WS-EDIT1 PIC 9(5).99 VALUE ZERO.
01 WS-EDIT2 PIC Z(3),ZZ9.99 VALUE ZERO.
"#,
        r#"
    MOVE WS-SRC TO WS-EDIT1.
    DISPLAY WS-EDIT1.
    MOVE WS-SRC TO WS-EDIT2.
    DISPLAY WS-EDIT2.
"#,
    ));
    assert_eq!(output, vec!["12345.67", " 12,345.67"]);
}
