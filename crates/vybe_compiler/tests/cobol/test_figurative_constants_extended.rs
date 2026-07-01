use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn figurative_constant_move_all_repeats_pattern() {
    let output = run_prints(&p(
        r#"
01 WS-LINE PIC X(6) VALUE SPACES.
"#,
        r#"
    MOVE ALL "-" TO WS-LINE.
    DISPLAY WS-LINE.
"#,
    ));
    assert_eq!(output, vec!["------"]);
}

#[test]
fn figurative_constant_move_all_to_numeric_field() {
    let output = run_prints(&p(
        r#"
01 WS-NUM PIC 9(4) VALUE 0.
"#,
        r#"
    MOVE ALL "9" TO WS-NUM.
    DISPLAY WS-NUM.
"#,
    ));
    assert_eq!(output, vec!["9999"]);
}

#[test]
fn figurative_constant_low_values_fill_field() {
    let output = run_prints(&p(
        r#"
01 WS-BUF PIC X(4) VALUE SPACES.
"#,
        r#"
    MOVE LOW-VALUES TO WS-BUF.
    DISPLAY WS-BUF.
"#,
    ));
    assert_eq!(output, vec!["\u{0}\u{0}\u{0}\u{0}"]);
}

#[test]
fn figurative_constant_high_values_fill_field() {
    let output = run_prints(&p(
        r#"
01 WS-BUF PIC X(4) VALUE SPACES.
"#,
        r#"
    MOVE HIGH-VALUES TO WS-BUF.
    DISPLAY WS-BUF.
"#,
    ));
    assert_eq!(output, vec!["\u{255}\u{255}\u{255}\u{255}"]);
}

#[test]
fn figurative_constant_zeros_assign_numeric_value() {
    let output = run_prints(&p(
        r#"
01 WS-N PIC 9(3) VALUE 123.
"#,
        r#"
    MOVE ZEROS TO WS-N.
    DISPLAY WS-N.
"#,
    ));
    assert_eq!(output, vec!["000"]);
}

#[test]
fn figurative_constant_spaces_assign_blank_value() {
    let output = run_prints(&p(
        r#"
01 WS-NAME PIC X(5) VALUE "HELLO".
"#,
        r#"
    MOVE SPACES TO WS-NAME.
    DISPLAY WS-NAME.
"#,
    ));
    assert_eq!(output, vec!["     "]);
}

#[test]
fn figurative_constant_compare_spaces_and_zeros() {
    compile_ok(&p(
        r#"
01 WS-NAME PIC X(5) VALUE SPACES.
01 WS-NUM PIC 9(3) VALUE ZEROS.
"#,
        r#"
    IF WS-NAME = SPACES
        DISPLAY "SPACE"
    END-IF.
    IF WS-NUM = ZEROS
        DISPLAY "ZERO"
    END-IF.
"#,
    ));
}

#[test]
fn figurative_constant_move_all_repeats_multiple_characters() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(8) VALUE SPACES.
"#,
        r#"
    MOVE ALL "12" TO WS-TXT.
    DISPLAY WS-TXT.
"#,
    ));
    assert_eq!(output, vec!["12121212"]);
}

#[test]
fn figurative_constant_high_value_is_comparable() {
    let output = run_prints(&p(
        r#"
01 WS-KEY PIC X(4) VALUE SPACES.
"#,
        r#"
    MOVE HIGH-VALUES TO WS-KEY.
    IF WS-KEY = HIGH-VALUES
        DISPLAY "HIGH"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["HIGH"]);
}

#[test]
fn figurative_constant_low_value_is_comparable() {
    let output = run_prints(&p(
        r#"
01 WS-KEY PIC X(4) VALUE SPACES.
"#,
        r#"
    MOVE LOW-VALUES TO WS-KEY.
    IF WS-KEY = LOW-VALUES
        DISPLAY "LOW"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["LOW"]);
}

#[test]
fn figurative_constant_zeroes_can_be_moved_into_group_item() {
    compile_ok(&p(
        r#"
01 WS-GRP.
   05 WS-A PIC 9(3) VALUE 7.
   05 WS-B PIC 9(3) VALUE 8.
"#,
        r#"
    MOVE ZEROS TO WS-GRP.
"#,
    ));
}

#[test]
fn figurative_constant_spaces_can_be_moved_into_group_item() {
    compile_ok(&p(
        r#"
01 WS-GRP.
   05 WS-A PIC X(2) VALUE "AA".
   05 WS-B PIC X(2) VALUE "BB".
"#,
        r#"
    MOVE SPACES TO WS-GRP.
"#,
    ));
}
