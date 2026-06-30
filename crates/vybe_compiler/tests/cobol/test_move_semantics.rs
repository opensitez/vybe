use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_move_multiple_destinations() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(5) VALUE "HELLO".
01 WS-DST1 PIC X(5) VALUE "AAAAA".
01 WS-DST2 PIC X(5) VALUE "BBBBB".
01 WS-DST3 PIC X(5) VALUE "CCCCC".
"#,
        r#"
    MOVE WS-SRC TO WS-DST1 WS-DST2 WS-DST3.
    DISPLAY WS-DST1.
    DISPLAY WS-DST2.
    DISPLAY WS-DST3.
"#,
    ));
    assert_eq!(output, vec!["HELLO", "HELLO", "HELLO"]);
}

#[test]
fn test_move_zeros_to_alphanumeric() {
    let output = run_prints(&p(
        "01 WS-DST PIC X(5) VALUE SPACES.",
        r#"
    MOVE ZEROS TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["00000"]);
}

#[test]
fn test_move_spaces_to_numeric() {
    let output = run_prints(&p(
        "01 WS-DST PIC 9(5) VALUE 12345.",
        r#"
    MOVE SPACES TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["00000"]);
}

#[test]
fn test_move_high_values_to_alpha() {
    compile_ok(&p(
        "01 WS-DST PIC X(5).",
        r#"
    MOVE HIGH-VALUES TO WS-DST.
"#,
    ));
}

#[test]
fn test_move_low_values_to_alpha() {
    compile_ok(&p(
        "01 WS-DST PIC X(5).",
        r#"
    MOVE LOW-VALUES TO WS-DST.
"#,
    ));
}

#[test]
fn test_move_all_literal_repeating() {
    let output = run_prints(&p(
        "01 WS-DST PIC X(6) VALUE SPACES.",
        r#"
    MOVE ALL "AB" TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["ABABAB"]);
}

#[test]
fn test_move_numeric_left_padding() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC 9(2) VALUE 42.
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
fn test_move_numeric_left_truncation() {
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
fn test_move_alpha_right_padding() {
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
fn test_move_alpha_right_truncation() {
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
fn test_move_decimal_to_integer() {
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
fn test_move_integer_to_decimal() {
    compile_ok(&p(
        r#"
01 WS-SRC PIC 9(3) VALUE 123.
01 WS-DST PIC 9(3)V99 VALUE 0.0.
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
"#,
    ));
}

#[test]
fn test_move_signed_to_unsigned() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC S9(3) VALUE -123.
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
fn test_move_numeric_string_to_numeric_field() {
    compile_ok(&p(
        r#"
01 WS-SRC PIC X(3) VALUE "123".
01 WS-DST PIC 9(3) VALUE 0.
"#,
        r#"
    MOVE WS-SRC TO WS-DST.
"#,
    ));
}

#[test]
fn test_move_group_to_elementary() {
    let output = run_prints(&p(
        r#"
01 WS-GROUP.
   05 WS-A PIC X(3) VALUE "ABC".
   05 WS-B PIC X(3) VALUE "DEF".
01 WS-DST PIC X(6) VALUE SPACES.
"#,
        r#"
    MOVE WS-GROUP TO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["ABCDEF"]);
}

#[test]
fn test_move_elementary_to_group() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(6) VALUE "ABCDEF".
01 WS-GROUP.
   05 WS-A PIC X(3) VALUE "XXX".
   05 WS-B PIC X(3) VALUE "YYY".
"#,
        r#"
    MOVE WS-SRC TO WS-GROUP.
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["ABC", "DEF"]);
}

#[test]
fn test_move_into_table_subscript() {
    let output = run_prints(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(2) OCCURS 5 TIMES.
"#,
        r#"
    MOVE 99 TO WS-ITEM(3).
    DISPLAY WS-ITEM(3).
"#,
    ));
    assert_eq!(output, vec!["99"]);
}

#[test]
fn test_move_into_refmod_target() {
    let output = run_prints(&p(
        "01 WS-TEXT PIC X(6) VALUE \"AABBCC\".",
        r#"
    MOVE "XX" TO WS-TEXT(3:2).
    DISPLAY WS-TEXT.
"#,
    ));
    assert_eq!(output, vec!["AAXXCC"]);
}

#[test]
fn test_move_quote_literal() {
    compile_ok(&p(
        "01 WS-DST PIC X(1).",
        r#"
    MOVE QUOTE TO WS-DST.
"#,
    ));
}
