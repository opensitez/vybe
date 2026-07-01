use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn perform_varying_with_step_and_limit() {
    let output = run_prints(&p(
        r#"
01 WS-I PIC 9 VALUE 0.
"#,
        r#"
    PERFORM VARYING WS-I FROM 1 BY 2 UNTIL WS-I > 5
        DISPLAY WS-I
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["1", "3", "5"]);
}

#[test]
fn perform_until_with_nested_condition() {
    let output = run_prints(&p(
        r#"
01 WS-I PIC 9 VALUE 0.
"#,
        r#"
    PERFORM UNTIL WS-I >= 3
        ADD 1 TO WS-I
        DISPLAY WS-I
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["1", "2", "3"]);
}

#[test]
fn evaluate_with_multiple_when_branches() {
    let output = run_prints(&p(
        r#"
01 WS-X PIC 9 VALUE 2.
"#,
        r#"
    EVALUATE WS-X
        WHEN 1
            DISPLAY "ONE"
        WHEN 2
            DISPLAY "TWO"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["TWO"]);
}

#[test]
fn evaluate_true_with_range_conditions() {
    let output = run_prints(&p(
        r#"
01 WS-AGE PIC 99 VALUE 22.
"#,
        r#"
    EVALUATE TRUE
        WHEN WS-AGE < 13
            DISPLAY "CHILD"
        WHEN WS-AGE < 20
            DISPLAY "TEEN"
        WHEN OTHER
            DISPLAY "ADULT"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["ADULT"]);
}

#[test]
fn evaluate_string_branching() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "B".
"#,
        r#"
    EVALUATE WS-CODE
        WHEN "A"
            DISPLAY "ALPHA"
        WHEN "B"
            DISPLAY "BETA"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["BETA"]);
}

#[test]
fn perform_paragraph_with_procedure_name() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM STEP-ONE.\n    STOP RUN.\nSTEP-ONE.\n    DISPLAY \"DONE\".",
    );
}

#[test]
fn perform_times_with_zero_iterations_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-I PIC 9 VALUE 0.
"#,
        r#"
    PERFORM 0 TIMES
        ADD 1 TO WS-I
    END-PERFORM.
"#,
    ));
}

#[test]
fn perform_varying_downward_reaches_limit() {
    let output = run_prints(&p(
        r#"
01 WS-I PIC 9 VALUE 5.
"#,
        r#"
    PERFORM VARYING WS-I FROM 5 BY -1 UNTIL WS-I < 3
        DISPLAY WS-I
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["5", "4", "3"]);
}

#[test]
fn evaluate_true_with_multiple_when_branches() {
    let output = run_prints(&p(
        r#"
01 WS-AGE PIC 99 VALUE 17.
"#,
        r#"
    EVALUATE TRUE
        WHEN WS-AGE < 13
            DISPLAY "CHILD"
        WHEN WS-AGE < 20
            DISPLAY "TEEN"
        WHEN OTHER
            DISPLAY "ADULT"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["TEEN"]);
}

#[test]
fn evaluate_with_string_and_numeric_alternatives() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "C".
"#,
        r#"
    EVALUATE WS-CODE
        WHEN "A"
            DISPLAY "ALPHA"
        WHEN "C"
            DISPLAY "CHARLIE"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["CHARLIE"]);
}

#[test]
fn evaluate_without_other_branch_is_accepted() {
    compile_ok(&p(
        r#"
01 WS-CODE PIC 9 VALUE 1.
"#,
        r#"
    EVALUATE WS-CODE
        WHEN 1
            DISPLAY "ONE"
        WHEN 2
            DISPLAY "TWO"
    END-EVALUATE.
"#,
    ));
}
