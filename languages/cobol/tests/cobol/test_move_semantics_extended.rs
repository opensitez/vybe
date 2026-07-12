use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn move_corresponding_between_groups_preserves_fields() {
    let output = run_prints(&p(
        r#"
01 WS-SRC.
   05 WS-A PIC X(3) VALUE "ABC".
   05 WS-B PIC X(3) VALUE "DEF".
01 WS-DST.
   05 WS-A PIC X(3) VALUE SPACES.
   05 WS-B PIC X(3) VALUE SPACES.
"#,
        r#"
    MOVE CORRESPONDING WS-SRC TO WS-DST.
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["ABC", "DEF"]);
}

#[test]
fn move_all_to_alphanumeric_field_repeats_pattern() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(6) VALUE SPACES.
"#,
        r#"
    MOVE ALL "AB" TO WS-TXT.
    DISPLAY WS-TXT.
"#,
    ));
    assert_eq!(output, vec!["ABABAB"]);
}

#[test]
fn move_all_to_group_field_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-GRP.
   05 WS-X PIC X(2) VALUE SPACES.
   05 WS-Y PIC X(2) VALUE SPACES.
"#,
        r#"
    MOVE ALL "Z" TO WS-GRP.
"#,
    ));
}

#[test]
fn move_numeric_zero_to_alphanumeric_field_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-TXT PIC X(5) VALUE SPACES.
"#,
        r#"
    MOVE ZEROS TO WS-TXT.
"#,
    ));
}

#[test]
fn move_spaces_to_numeric_field_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-NUM PIC 9(3) VALUE 123.
"#,
        r#"
    MOVE SPACES TO WS-NUM.
"#,
    ));
}

#[test]
fn move_high_values_to_alphanumeric_field_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-TXT PIC X(5) VALUE SPACES.
"#,
        r#"
    MOVE HIGH-VALUES TO WS-TXT.
"#,
    ));
}

#[test]
fn move_low_values_to_alphanumeric_field_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-TXT PIC X(5) VALUE SPACES.
"#,
        r#"
    MOVE LOW-VALUES TO WS-TXT.
"#,
    ));
}

#[test]
fn move_all_to_group_item_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-GRP.
   05 WS-A PIC X(2) VALUE SPACES.
   05 WS-B PIC X(2) VALUE SPACES.
"#,
        r#"
    MOVE ALL "X" TO WS-GRP.
"#,
    ));
}

#[test]
fn move_zeroes_to_group_item_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-GRP.
   05 WS-A PIC 9(2) VALUE 1.
   05 WS-B PIC 9(2) VALUE 2.
"#,
        r#"
    MOVE ZEROES TO WS-GRP.
"#,
    ));
}

#[test]
fn move_spaces_to_group_item_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-GRP.
   05 WS-A PIC X(2) VALUE "A".
   05 WS-B PIC X(2) VALUE "B".
"#,
        r#"
    MOVE SPACES TO WS-GRP.
"#,
    ));
}
