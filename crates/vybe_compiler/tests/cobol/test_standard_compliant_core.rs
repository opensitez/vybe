use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn perform_varying_accumulates_expected_total() {
    let output = run_prints(&p(
        r#"
01 WS-I PIC 9 VALUE 1.
01 WS-SUM PIC 9(3) VALUE 0.
"#,
        r#"
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5
        ADD WS-I TO WS-SUM
    END-PERFORM.
    DISPLAY WS-SUM.
"#,
    ));
    assert_eq!(output, vec!["15"]);
}

#[test]
fn evaluate_true_selects_correct_grade_band() {
    let output = run_prints(&p(
        r#"
01 WS-SCORE PIC 9(3) VALUE 82.
01 WS-GRADE PIC X VALUE "?".
"#,
        r#"
    EVALUATE TRUE
        WHEN WS-SCORE >= 90
            MOVE "A" TO WS-GRADE
        WHEN WS-SCORE >= 80
            MOVE "B" TO WS-GRADE
        WHEN OTHER
            MOVE "C" TO WS-GRADE
    END-EVALUATE.
    DISPLAY WS-GRADE.
"#,
    ));
    assert_eq!(output, vec!["B"]);
}

#[test]
fn add_giving_produces_expected_result() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(2) VALUE 7.
01 WS-B PIC 9(2) VALUE 8.
01 WS-R PIC 9(3) VALUE 0.
"#,
        r#"
    ADD WS-A WS-B GIVING WS-R.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["15"]);
}

#[test]
fn string_and_unstring_round_trip_fields() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(3) VALUE "ONE".
01 WS-B PIC X(3) VALUE "TWO".
01 WS-COMBINED PIC X(8) VALUE SPACES.
01 WS-R1 PIC X(3) VALUE SPACES.
01 WS-R2 PIC X(3) VALUE SPACES.
"#,
        r#"
    STRING WS-A DELIMITED BY SIZE
           "," DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           INTO WS-COMBINED.
    UNSTRING WS-COMBINED DELIMITED BY "," INTO WS-R1 WS-R2.
    DISPLAY WS-R1.
    DISPLAY WS-R2.
"#,
    ));
    assert_eq!(output, vec!["ONE", "TWO"]);
}

#[test]
fn inspect_tallying_and_replacing_works_on_same_field() {
    let output = run_prints(&p(
        r#"
01 WS-TEXT PIC X(8) VALUE "ABABXABA".
01 WS-COUNT PIC 9 VALUE 0.
"#,
        r#"
    INSPECT WS-TEXT TALLYING WS-COUNT FOR ALL "A".
    DISPLAY WS-COUNT.
"#,
    ));
    assert_eq!(output, vec!["4"]);
}

#[test]
fn string_with_pointer_updates_destination_and_pointer() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(2) VALUE "AB".
01 WS-B PIC X(2) VALUE "CD".
01 WS-R PIC X(8) VALUE SPACES.
"#,
        r#"
    STRING WS-A DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           INTO WS-R.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["ABCD"]);
}
