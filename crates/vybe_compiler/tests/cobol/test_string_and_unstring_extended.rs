use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn string_with_multiple_sources_and_delimiters() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(3) VALUE "A".
01 WS-B PIC X(3) VALUE "B".
01 WS-C PIC X(3) VALUE "C".
01 WS-R PIC X(12) VALUE SPACES.
"#,
        r#"
    STRING WS-A DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           WS-C DELIMITED BY SIZE
           INTO WS-R.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["ABC"]);
}

#[test]
fn string_with_literal_and_variable_sources() {
    let output = run_prints(&p(
        r#"
01 WS-NAME PIC X(5) VALUE "COBOL".
01 WS-R PIC X(12) VALUE SPACES.
"#,
        r#"
    STRING "HELLO " DELIMITED BY SIZE
           WS-NAME DELIMITED BY SIZE
           INTO WS-R.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["HELLO COBOL"]);
}

#[test]
fn unstring_with_multiple_targets_and_delimiters() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(12) VALUE "A,B,C".
01 WS-F1 PIC X(3) VALUE SPACES.
01 WS-F2 PIC X(3) VALUE SPACES.
01 WS-F3 PIC X(3) VALUE SPACES.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY "," INTO WS-F1 WS-F2 WS-F3.
    DISPLAY WS-F1.
    DISPLAY WS-F2.
    DISPLAY WS-F3.
"#,
    ));
    assert_eq!(output, vec!["A", "B", "C"]);
}

#[test]
fn unstring_with_space_delimiter() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(12) VALUE "ONE TWO".
01 WS-F1 PIC X(5) VALUE SPACES.
01 WS-F2 PIC X(5) VALUE SPACES.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY SPACE INTO WS-F1 WS-F2.
    DISPLAY WS-F1.
    DISPLAY WS-F2.
"#,
    ));
    assert_eq!(output, vec!["ONE", "TWO"]);
}

#[test]
fn inspect_tallying_with_multiple_patterns() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(12) VALUE "ABBAABBA".
01 WS-CNT PIC 9(3) VALUE 0.
"#,
        r#"
    INSPECT WS-TXT TALLYING WS-CNT FOR ALL "A".
    DISPLAY WS-CNT.
"#,
    ));
    assert_eq!(output, vec!["4"]);
}

#[test]
fn inspect_replacing_with_characters() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(8) VALUE "ABC123".
"#,
        r#"
    INSPECT WS-TXT REPLACING CHARACTERS BY "X".
    DISPLAY WS-TXT.
"#,
    ));
    assert_eq!(output, vec!["XXXXXXX"]);
}

#[test]
fn inspect_converting_letters_to_upper_case() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(8) VALUE "abc123".
"#,
        r#"
    INSPECT WS-TXT CONVERTING "abc" TO "ABC".
    DISPLAY WS-TXT.
"#,
    ));
    assert_eq!(output, vec!["ABC123"]);
}

#[test]
fn reference_modification_on_string_field() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(10) VALUE "HELLOTEST".
01 WS-SUB PIC X(5) VALUE SPACES.
"#,
        r#"
    MOVE WS-TXT(1:5) TO WS-SUB.
    DISPLAY WS-SUB.
"#,
    ));
    assert_eq!(output, vec!["HELLO"]);
}

#[test]
fn string_with_space_delimiter_concatenates_fields() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(4) VALUE "ONE".
01 WS-B PIC X(4) VALUE "TWO".
01 WS-R PIC X(20) VALUE SPACES.
"#,
        r#"
    STRING WS-A DELIMITED BY SPACE
           WS-B DELIMITED BY SPACE
           INTO WS-R.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["ONETWO"]);
}

#[test]
fn unstring_with_all_delimiter_splits_repeated_separators() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(12) VALUE "A,,B".
01 WS-F1 PIC X(3) VALUE SPACES.
01 WS-F2 PIC X(3) VALUE SPACES.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY ALL "," INTO WS-F1 WS-F2.
    DISPLAY WS-F1.
    DISPLAY WS-F2.
"#,
    ));
    assert_eq!(output, vec!["A", "B"]);
}

#[test]
fn inspect_tallying_for_leading_zeroes_counts_prefix() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(8) VALUE "0001234".
01 WS-CNT PIC 9(3) VALUE 0.
"#,
        r#"
    INSPECT WS-TXT TALLYING WS-CNT FOR LEADING "0".
    DISPLAY WS-CNT.
"#,
    ));
    assert_eq!(output, vec!["3"]);
}

#[test]
fn inspect_replacing_first_character_changes_only_first_match() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(6) VALUE "AAAAAA".
"#,
        r#"
    INSPECT WS-TXT REPLACING FIRST "A" BY "B".
    DISPLAY WS-TXT.
"#,
    ));
    assert_eq!(output, vec!["BAAAAA"]);
}

#[test]
fn string_with_pointer_updates_destination_and_pointer() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(2) VALUE "AB".
01 WS-B PIC X(2) VALUE "CD".
01 WS-R PIC X(8) VALUE SPACES.
01 WS-PTR PIC 9(2) VALUE 1.
"#,
        r#"
    STRING WS-A DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           INTO WS-R
           WITH POINTER WS-PTR.
    DISPLAY WS-R.
    DISPLAY WS-PTR.
"#,
    ));
    assert_eq!(output, vec!["ABCD", "5"]);
}

#[test]
fn unstring_with_count_in_reports_token_lengths() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(12) VALUE "AA,BBB".
01 WS-F1 PIC X(5) VALUE SPACES.
01 WS-F2 PIC X(5) VALUE SPACES.
01 WS-C1 PIC 9(2) VALUE 0.
01 WS-C2 PIC 9(2) VALUE 0.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY ","
        INTO WS-F1 COUNT IN WS-C1
             WS-F2 COUNT IN WS-C2.
    DISPLAY WS-F1.
    DISPLAY WS-F2.
    DISPLAY WS-C1.
    DISPLAY WS-C2.
"#,
    ));
    assert_eq!(output, vec!["AA", "BBB", "2", "3"]);
}

#[test]
fn unstring_tallying_in_counts_receivers_used() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(12) VALUE "A,B,C".
01 WS-F1 PIC X(2) VALUE SPACES.
01 WS-F2 PIC X(2) VALUE SPACES.
01 WS-F3 PIC X(2) VALUE SPACES.
01 WS-T PIC 9 VALUE 0.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY ","
        INTO WS-F1 WS-F2 WS-F3
        TALLYING IN WS-T.
    DISPLAY WS-F1.
    DISPLAY WS-F2.
    DISPLAY WS-F3.
    DISPLAY WS-T.
"#,
    ));
    assert_eq!(output, vec!["A", "B", "C", "3"]);
}

#[test]
fn string_on_overflow_executes_overflow_branch() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(3) VALUE "ABC".
01 WS-B PIC X(3) VALUE "DEF".
01 WS-R PIC X(3) VALUE SPACES.
"#,
        r#"
    STRING WS-A DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           INTO WS-R
      ON OVERFLOW DISPLAY "OVF"
      NOT ON OVERFLOW DISPLAY "OK"
    END-STRING.
"#,
    ));
    assert_eq!(output, vec!["OVF"]);
}
