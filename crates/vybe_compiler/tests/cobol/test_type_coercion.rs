use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_coercion_numeric_to_alpha() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(2) VALUE 42.
01 WS-DST PIC X(5) VALUE SPACES.
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    // Converts to character display representation: "42   "
    assert_eq!(output, vec!["42   "]);
}

#[test]
fn test_coercion_numeric_truncation() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(5) VALUE 12345.
01 WS-DST PIC 9(3) VALUE 0.
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["345"]);
}

#[test]
fn test_coercion_numeric_padding() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(3) VALUE 42.
01 WS-DST PIC 9(5) VALUE 0.
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["00042"]);
}

#[test]
fn test_coercion_alpha_padding() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(3) VALUE "ABC".
01 WS-DST PIC X(6) VALUE "XXXXXX".
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["ABC   "]);
}

#[test]
fn test_coercion_alpha_truncation() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(6) VALUE "ABCDEF".
01 WS-DST PIC X(3) VALUE "XXX".
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["ABC"]);
}

#[test]
fn test_coercion_decimal_to_integer() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(3)V99 VALUE 123.45.
01 WS-DST PIC 9(3) VALUE 0.
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["123"]);
}

#[test]
fn test_coercion_zeros_to_alpha() {
    let output = run_prints(&p(
        "01 WS-DST PIC X(5) VALUE SPACES.",
        r#"
    MOVE ZEROS TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["00000"]);
}
