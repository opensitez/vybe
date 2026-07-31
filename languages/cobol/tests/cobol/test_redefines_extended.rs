use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn redefines_allows_alphanumeric_and_numeric_views() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(6).
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
"#,
        r#"
    MOVE 123456 TO WS-NUMBER.
    DISPLAY WS-NUMBER.
    DISPLAY WS-BUFFER.
"#,
    ));
    assert_eq!(output, vec!["123456", "123456"]);
}

#[test]
fn redefines_preserves_storage_when_moving_data() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(4) VALUE "ABCD".
01 WS-HEX REDEFINES WS-BUFFER PIC X(4).
"#,
        r#"
    DISPLAY WS-HEX.
"#,
    ));
    assert_eq!(output, vec!["ABCD"]);
}

#[test]
fn redefines_with_group_item_is_accepted() {
    let output = run_prints(&p(
        r#"
01 WS-REC.
   05 WS-FIELD1 PIC X(2) VALUE "AA".
   05 WS-FIELD2 PIC X(2) VALUE "BB".
01 WS-ALT REDEFINES WS-REC.
   05 WS-VAL PIC X(4).
"#,
        r#"
    DISPLAY WS-VAL.
"#,
    ));
    assert_eq!(output, vec!["AABB"]);
}

#[test]
fn redefines_can_be_used_with_move_to_same_storage() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(6) VALUE SPACES.
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
"#,
        r#"
    MOVE 42 TO WS-NUMBER.
    DISPLAY WS-BUFFER.
"#,
    ));
    assert_eq!(output, vec!["000042"]);
}

#[test]
fn redefines_can_shadow_alphanumeric_value_with_numeric_view() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(4) VALUE "1234".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(4).
"#,
        r#"
    MOVE 9999 TO WS-NUMBER.
    DISPLAY WS-BUFFER.
"#,
    ));
    assert_eq!(output, vec!["9999"]);
}

#[test]
fn redefines_with_group_item_can_be_displayed() {
    let output = run_prints(&p(
        r#"
01 WS-REC.
   05 WS-A PIC X(2) VALUE "AB".
   05 WS-B PIC X(2) VALUE "CD".
01 WS-ALT REDEFINES WS-REC.
   05 WS-VAL PIC X(4).
"#,
        r#"
    DISPLAY WS-VAL.
"#,
    ));
    assert_eq!(output, vec!["ABCD"]);
}

#[test]
fn redefines_with_spaces_as_numeric() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(6) VALUE "ABC123".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
01 WS-FLAG PIC X VALUE 'N'.
"#,
        r#"
    MOVE SPACES TO WS-NUMBER.
    IF WS-NUMBER = 0
        MOVE 'Y' TO WS-FLAG
    END-IF.
    DISPLAY WS-FLAG.
"#,
    ));
    assert_eq!(output, vec!["Y"]);
}

#[test]
fn redefines_with_zeros_as_numeric() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(6) VALUE "ABC123".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
01 WS-FLAG PIC X VALUE 'N'.
"#,
        r#"
    MOVE ZEROS TO WS-NUMBER.
    IF WS-NUMBER = 0
        MOVE 'Y' TO WS-FLAG
    END-IF.
    DISPLAY WS-FLAG.
"#,
    ));
    assert_eq!(output, vec!["Y"]);
}

#[test]
fn redefines_numeric_and_alpha_roundtrip() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(4) VALUE "0000".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(4).
01 WS-NEXT PIC X(2) VALUE "AA".
"#,
        r#"
    MOVE 77 TO WS-NUMBER.
    DISPLAY WS-BUFFER.
    MOVE WS-BUFFER TO WS-NEXT.
    DISPLAY WS-NEXT.
"#,
    ));
    assert_eq!(output, vec!["0077", "00"]);
}

#[test]
fn redefines_multiple_nested_views() {
    let output = run_prints(&p(
        r#"
01 WS-BUFFER PIC X(8) VALUE "12345678".
01 WS-OUTER REDEFINES WS-BUFFER.
   05 WS-HI PIC 9(4).
   05 WS-LO PIC 9(4).
"#,
        r#"
    ADD 1 TO WS-HI.
    SUBTRACT 1 FROM WS-LO.
    DISPLAY WS-HI.
    DISPLAY WS-LO.
"#,
    ));
    assert_eq!(output, vec!["1235", "5677"]);
}
