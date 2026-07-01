use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn redefines_allows_alphanumeric_and_numeric_views() {
    compile_ok(&p(
        r#"
01 WS-BUFFER PIC X(6).
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
"#,
        r#"
    MOVE 123456 TO WS-NUMBER.
"#,
    ));
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
    compile_ok(&p(
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
}

#[test]
fn redefines_can_be_used_with_move_to_same_storage() {
    compile_ok(&p(
        r#"
01 WS-BUFFER PIC X(6) VALUE SPACES.
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
"#,
        r#"
    MOVE 42 TO WS-NUMBER.
"#,
    ));
}

#[test]
fn redefines_can_shadow_alphanumeric_value_with_numeric_view() {
    compile_ok(&p(
        r#"
01 WS-BUFFER PIC X(4) VALUE "1234".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(4).
"#,
        r#"
    MOVE 9999 TO WS-NUMBER.
"#,
    ));
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
fn redefines_can_be_used_with_move_spaces() {
    compile_ok(&p(
        r#"
01 WS-BUFFER PIC X(6) VALUE "ABC123".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
"#,
        r#"
    MOVE SPACES TO WS-NUMBER.
"#,
    ));
}

#[test]
fn redefines_can_be_used_with_move_zeros() {
    compile_ok(&p(
        r#"
01 WS-BUFFER PIC X(6) VALUE "ABC123".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(6).
"#,
        r#"
    MOVE ZEROS TO WS-NUMBER.
"#,
    ));
}
