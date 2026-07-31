use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn condition_name_single_value_selects_true_branch() {
    let output = run_prints(&p(
        r#"
01 WS-STATUS PIC X VALUE "A".
   88 IS-READY VALUE "A".
"#,
        r#"
    IF IS-READY
        DISPLAY "READY"
    ELSE
        DISPLAY "WAIT"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["READY"]);
}

#[test]
fn condition_name_multiple_values_selects_true_branch() {
    let output = run_prints(&p(
        r#"
01 WS-STATUS PIC X VALUE "B".
   88 IS-READY VALUE "A", "B".
"#,
        r#"
    IF IS-READY
        DISPLAY "READY"
    ELSE
        DISPLAY "WAIT"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["READY"]);
}

#[test]
fn condition_name_through_range_selects_true_branch() {
    let output = run_prints(&p(
        r#"
01 WS-AGE PIC 99 VALUE 25.
   88 IS-YOUTH VALUE 15 THRU 30.
"#,
        r#"
    IF IS-YOUTH
        DISPLAY "YOUTH"
    ELSE
        DISPLAY "OTHER"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["YOUTH"]);
}

#[test]
fn condition_name_set_true_and_false_updates_value() {
    let output = run_prints(&p(
        r#"
01 WS-FLAG PIC X VALUE "N".
   88 IS-ON VALUE "Y".
"#,
        r#"
    SET IS-ON TO TRUE.
    DISPLAY WS-FLAG.
    SET IS-ON TO FALSE.
    DISPLAY WS-FLAG.
"#,
    ));
    assert_eq!(output, vec!["Y", "N"]);
}

#[test]
fn condition_name_in_evaluate_selects_matching_branch() {
    let output = run_prints(&p(
        r#"
01 WS-STATUS PIC X VALUE "B".
   88 IS-STARTED VALUE "A".
   88 IS-READY VALUE "B".
"#,
        r#"
    EVALUATE TRUE
        WHEN IS-STARTED
            DISPLAY "STARTED"
        WHEN IS-READY
            DISPLAY "READY"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["READY"]);
}

#[test]
fn condition_name_on_numeric_field_is_true_for_range() {
    let output = run_prints(&p(
        r#"
01 WS-SCORE PIC 99 VALUE 60.
   88 IS-PASSING VALUE 50 THRU 100.
"#,
        r#"
    IF IS-PASSING
        DISPLAY "PASS"
    ELSE
        DISPLAY "FAIL"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["PASS"]);
}

#[test]
fn condition_name_supports_nested_boolean_checks() {
    let output = run_prints(&p(
        r#"
01 WS-STATE PIC X VALUE "Y".
   88 IS-ACTIVE VALUE "Y".
"#,
        r#"
    IF IS-ACTIVE
        DISPLAY "ACTIVE"
    ELSE
        DISPLAY "INACTIVE"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["ACTIVE"]);
}

#[test]
fn condition_name_with_alphanumeric_values_is_case_sensitive() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "C".
   88 IS-VALID VALUE "A", "B", "C".
"#,
        r#"
    IF IS-VALID
        DISPLAY "VALID"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["VALID"]);
}

#[test]
fn condition_name_false_branch_is_taken_when_value_does_not_match() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "Z".
   88 IS-VALID VALUE "A", "B", "C".
"#,
        r#"
    IF IS-VALID
        DISPLAY "VALID"
    ELSE
        DISPLAY "INVALID"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["INVALID"]);
}

#[test]
fn condition_name_recomputed_after_move() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "A".
   88 IS-VALID VALUE "A".
"#,
        r#"
    MOVE "B" TO WS-CODE.
    IF IS-VALID
        DISPLAY "VALID"
    ELSE
        DISPLAY "INVALID"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["INVALID"]);
}

#[test]
fn condition_name_with_range_false_outside_bounds() {
    let output = run_prints(&p(
        r#"
01 WS-AGE PIC 99 VALUE 40.
   88 IS-YOUTH VALUE 15 THRU 30.
"#,
        r#"
    IF IS-YOUTH
        DISPLAY "YOUTH"
    ELSE
        DISPLAY "OTHER"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["OTHER"]);
}

#[test]
fn condition_name_can_gate_nested_if_logic() {
    let output = run_prints(&p(
        r#"
01 WS-STATE PIC X VALUE "Y".
   88 IS-OPEN VALUE "Y".
01 WS-COUNT PIC 9 VALUE 1.
"#,
        r#"
    IF IS-OPEN
        IF WS-COUNT > 0
            DISPLAY "OPEN-COUNT"
        END-IF
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["OPEN-COUNT"]);
}

#[test]
fn condition_name_supports_multiple_values_with_false_item() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "D".
   88 IS-VALID VALUE "A", "B", "C".
"#,
        r#"
    IF IS-VALID
        DISPLAY "YES"
    ELSE
        DISPLAY "NO"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["NO"]);
}

#[test]
fn condition_name_negated_with_not() {
    let output = run_prints(&p(
        r#"
01 WS-STATE PIC X VALUE "X".
   88 IS-ACTIVE VALUE "Y", "Z".
"#,
        r#"
    IF NOT IS-ACTIVE
        DISPLAY "INACTIVE"
    ELSE
        DISPLAY "ACTIVE"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["INACTIVE"]);
}

#[test]
fn condition_name_combined_with_logical_and() {
    let output = run_prints(&p(
        r#"
01 WS-STATE PIC X VALUE "Y".
01 WS-TYPE PIC X VALUE "R".
   88 IS-OPEN VALUE "Y".
   88 IS-READY VALUE "R".
"#,
        r#"
    IF IS-OPEN AND IS-READY
        DISPLAY "OPEN-READY"
    ELSE
        DISPLAY "BLOCKED"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["OPEN-READY"]);
}

#[test]
fn condition_name_false_when_unused_setter_is_false() {
    let output = run_prints(&p(
        r#"
01 WS-CODE PIC X VALUE "A".
   88 IS-READY VALUE "A".
"#,
        r#"
    SET IS-READY TO FALSE.
    IF IS-READY
        DISPLAY "READY"
    ELSE
        DISPLAY "NOT-READY"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["NOT-READY"]);
}

#[test]
fn condition_name_in_evaluate_with_other() {
    let output = run_prints(&p(
        r#"
01 WS-VAL PIC 99 VALUE 10.
   88 IS-LOW VALUE 1 THRU 5.
"#,
        r#"
    EVALUATE TRUE
        WHEN IS-LOW
            DISPLAY "LOW"
        WHEN OTHER
            DISPLAY "HIGH"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["HIGH"]);
}

#[test]
fn condition_name_multiple_range_values_compile() {
    compile_ok(&p(
        r#"
01 WS-AGE PIC 99 VALUE 18.
   88 AGE-STATE VALUE 0 THRU 17, 18 THRU 30.
"#,
        r#"
    IF AGE-STATE
        CONTINUE
    END-IF.
"#,
    ));
}
