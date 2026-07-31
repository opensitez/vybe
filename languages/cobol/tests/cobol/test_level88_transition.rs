use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn level88_basic_condition_true() {
    let out = run_prints(&p(
        "01 STATUS PIC X VALUE \"A\".\n    88 IS-ACTIVE VALUE \"A\".",
        "    IF IS-ACTIVE\n        DISPLAY \"ACTIVE\"\n    ELSE\n        DISPLAY \"INACTIVE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ACTIVE"]);
}

#[test]
fn level88_condition_false() {
    let out = run_prints(&p(
        "01 STATUS PIC X VALUE \"B\".\n    88 IS-ACTIVE VALUE \"A\".",
        "    IF IS-ACTIVE\n        DISPLAY \"ACTIVE\"\n    ELSE\n        DISPLAY \"INACTIVE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["INACTIVE"]);
}

#[test]
fn level88_set_to_true() {
    let out = run_prints(&p(
        "01 FLAG PIC X VALUE \"N\".\n    88 FLAG-ON VALUE \"Y\".\n    88 FLAG-OFF VALUE \"N\".",
        "    SET FLAG-ON TO TRUE.\n    IF FLAG-ON\n        DISPLAY \"ON\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ON"]);
}

#[test]
fn level88_set_to_false() {
    compile_ok(&p(
        "01 FLAG PIC X VALUE \"Y\".\n    88 FLAG-ON VALUE \"Y\".\n    88 FLAG-OFF VALUE \"N\".",
        "    SET FLAG-ON TO FALSE.",
    ));
}

#[test]
fn level88_multiple_values() {
    let out = run_prints(&p(
        "01 GRADE PIC X VALUE \"B\".\n    88 PASSING VALUE \"A\" \"B\" \"C\".",
        "    IF PASSING\n        DISPLAY \"PASS\"\n    ELSE\n        DISPLAY \"FAIL\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["PASS"]);
}

#[test]
fn level88_range_value() {
    compile_ok(&p(
        "01 SCORE PIC 9(3) VALUE 75.\n    88 PASSING-SCORE VALUE 60 THRU 100.",
        "    IF PASSING-SCORE\n        DISPLAY \"PASS\"\n    END-IF.",
    ));
}

#[test]
fn level88_in_evaluate_when() {
    let out = run_prints(&p(
        "01 STATUS PIC X VALUE \"A\".\n    88 IS-OPEN VALUE \"A\".\n    88 IS-CLOSED VALUE \"C\".",
        "    EVALUATE TRUE\n        WHEN IS-OPEN\n            DISPLAY \"OPEN\"\n        WHEN IS-CLOSED\n            DISPLAY \"CLOSED\"\n        WHEN OTHER\n            DISPLAY \"UNKNOWN\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["OPEN"]);
}

#[test]
fn level88_in_perform_until() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.\n    88 DONE VALUE 5.",
        "    PERFORM UNTIL DONE\n        ADD 1 TO N\n    END-PERFORM.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn level88_not_condition() {
    let out = run_prints(&p(
        "01 FLAG PIC X VALUE \"N\".\n    88 ACTIVE VALUE \"Y\".",
        "    IF NOT ACTIVE\n        DISPLAY \"INACTIVE\"\n    ELSE\n        DISPLAY \"ACTIVE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["INACTIVE"]);
}

#[test]
fn level88_numeric_values() {
    let out = run_prints(&p(
        "01 CODE PIC 9 VALUE 3.\n    88 SUCCESS VALUE 0.\n    88 FAILURE VALUE 1 THRU 9.",
        "    IF FAILURE\n        DISPLAY \"ERROR\"\n    ELSE\n        DISPLAY \"OK\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ERROR"]);
}

#[test]
fn level88_two_conditions_on_one_field() {
    let out = run_prints(&p(
        "01 COLOR PIC X VALUE \"R\".\n    88 RED VALUE \"R\".\n    88 BLUE VALUE \"B\".",
        "    EVALUATE TRUE\n        WHEN RED DISPLAY \"RED\"\n        WHEN BLUE DISPLAY \"BLUE\"\n        WHEN OTHER DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["RED"]);
}

#[test]
fn level88_set_true_then_test_field_value() {
    let out = run_prints(&p(
        "01 FLAG PIC X VALUE \"N\".\n    88 YES-FLAG VALUE \"Y\".",
        "    SET YES-FLAG TO TRUE.\n    DISPLAY FLAG.",
    ));
    assert_eq!(out, vec!["Y"]);
}

#[test]
fn level88_group_field_condition() {
    compile_ok(&p(
        "01 RECORD-STATUS PIC X VALUE \"A\".\n    88 RECORD-ACTIVE VALUE \"A\".\n    88 RECORD-DELETED VALUE \"D\".",
        "    IF RECORD-ACTIVE\n        DISPLAY \"ACTIVE\"\n    END-IF.",
    ));
}

#[test]
fn level88_in_and_condition() {
    let out = run_prints(&p(
        "01 STATUS PIC X VALUE \"Y\".\n    88 ENABLED VALUE \"Y\".\n01 COUNT PIC 9 VALUE 5.",
        "    IF ENABLED AND COUNT > 0\n        DISPLAY \"ACTIVE-AND-NONZERO\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ACTIVE-AND-NONZERO"]);
}

#[test]
fn level88_or_with_other_condition() {
    let out = run_prints(&p(
        "01 STATUS PIC X VALUE \"B\".\n    88 SPECIAL VALUE \"A\".\n01 N PIC 9 VALUE 5.",
        "    IF SPECIAL OR N > 3\n        DISPLAY \"EITHER\"\n    ELSE\n        DISPLAY \"NEITHER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["EITHER"]);
}

#[test]
fn level88_three_state_flag() {
    let out = run_prints(&p(
        r#"01 TRAFFIC PIC X VALUE "G".
    88 GREEN-LIGHT VALUE "G".
    88 YELLOW-LIGHT VALUE "Y".
    88 RED-LIGHT VALUE "R"."#,
        "    IF GREEN-LIGHT\n        DISPLAY \"GO\"\n    ELSE IF YELLOW-LIGHT\n        DISPLAY \"SLOW\"\n    ELSE\n        DISPLAY \"STOP\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["GO"]);
}

#[test]
fn level88_space_value_condition() {
    let out = run_prints(&p(
        "01 S PIC X VALUE SPACE.\n    88 IS-BLANK VALUE SPACE.",
        "    IF IS-BLANK\n        DISPLAY \"BLANK\"\n    ELSE\n        DISPLAY \"NOT BLANK\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["BLANK"]);
}

#[test]
fn level88_zero_numeric_condition() {
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE 0.\n    88 IS-ZERO VALUE 0.",
        "    IF IS-ZERO\n        DISPLAY \"ZERO\"\n    ELSE\n        DISPLAY \"NONZERO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn level88_field_after_set_can_be_tested_multiple_times() {
    let out = run_prints(&p(
        "01 F PIC X VALUE \"N\".\n    88 YES VALUE \"Y\".\n    88 NO-FLAG VALUE \"N\".",
        "    IF NO-FLAG\n        SET YES TO TRUE\n    END-IF.\n    IF YES\n        DISPLAY \"YES\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn level88_in_loop_termination() {
    let out = run_prints(&p(
        "01 COUNT PIC 9(3) VALUE 0.\n    88 LIMIT-REACHED VALUE 100.",
        "    PERFORM UNTIL LIMIT-REACHED\n        ADD 1 TO COUNT\n    END-PERFORM.\n    DISPLAY COUNT.",
    ));
    assert_eq!(out, vec!["100"]);
}

#[test]
fn level88_at_boundary_inclusive() {
    compile_ok(&p(
        "01 TEMP PIC S9(3) VALUE -5.\n    88 FREEZING VALUE -50 THRU 0.",
        "    IF FREEZING\n        DISPLAY \"COLD\"\n    END-IF.",
    ));
}

#[test]
fn level88_alphabetic_code() {
    let out = run_prints(&p(
        "01 CODE PIC X(2) VALUE \"AB\".\n    88 VALID-CODE VALUE \"AB\" \"CD\" \"EF\".",
        "    IF VALID-CODE\n        DISPLAY \"VALID\"\n    ELSE\n        DISPLAY \"INVALID\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["VALID"]);
}

#[test]
fn level88_invalid_code() {
    let out = run_prints(&p(
        "01 CODE PIC X(2) VALUE \"ZZ\".\n    88 VALID-CODE VALUE \"AB\" \"CD\" \"EF\".",
        "    IF VALID-CODE\n        DISPLAY \"VALID\"\n    ELSE\n        DISPLAY \"INVALID\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["INVALID"]);
}

#[test]
fn level88_set_false_by_first_false_value() {
    compile_ok(&p(
        "01 SWITCH PIC X VALUE \"Y\".\n    88 SW-ON VALUE \"Y\".\n    88 SW-OFF VALUE \"N\".",
        "    SET SW-ON TO FALSE.",
    ));
}

#[test]
fn level88_nested_if_with_multiple_flags() {
    let out = run_prints(&p(
        "01 A-FLAG PIC X VALUE \"Y\".\n    88 A-ON VALUE \"Y\".\n01 B-FLAG PIC X VALUE \"N\".\n    88 B-ON VALUE \"Y\".",
        "    IF A-ON\n        IF B-ON\n            DISPLAY \"BOTH\"\n        ELSE\n            DISPLAY \"ONLY A\"\n        END-IF\n    ELSE\n        DISPLAY \"NOT A\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ONLY A"]);
}

#[test]
fn level88_reuses_after_move() {
    let out = run_prints(&p(
        "01 S PIC X VALUE \"N\".\n    88 FLAGGED VALUE \"Y\".",
        "    MOVE \"Y\" TO S.\n    IF FLAGGED\n        DISPLAY \"SET\"\n    ELSE\n        DISPLAY \"UNSET\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["SET"]);
}

#[test]
fn level88_evaluate_true_multiple_when() {
    let out = run_prints(&p(
        "01 CODE PIC 9 VALUE 2.\n    88 CODE-ONE VALUE 1.\n    88 CODE-TWO VALUE 2.\n    88 CODE-THREE VALUE 3.",
        "    EVALUATE TRUE\n        WHEN CODE-ONE\n            DISPLAY \"ONE\"\n        WHEN CODE-TWO\n            DISPLAY \"TWO\"\n        WHEN CODE-THREE\n            DISPLAY \"THREE\"\n        WHEN OTHER\n            DISPLAY \"MANY\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["TWO"]);
}

#[test]
fn level88_boolean_sentinel_in_process_loop() {
    let out = run_prints(&p(
        "01 MORE-DATA PIC X VALUE \"Y\".\n    88 HAVE-DATA VALUE \"Y\".\n01 C PIC 9 VALUE 0.",
        "    PERFORM UNTIL NOT HAVE-DATA\n        ADD 1 TO C\n        IF C >= 3\n            MOVE \"N\" TO MORE-DATA\n        END-IF\n    END-PERFORM.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn level88_in_compute_guard() {
    let out = run_prints(&p(
        "01 DENOM PIC 9 VALUE 0.\n    88 NONZERO VALUE 1 THRU 9.\n01 R PIC 9(3) VALUE 0.",
        "    IF NONZERO\n        COMPUTE R = 100 / DENOM\n    ELSE\n        DISPLAY \"DIV BY ZERO GUARDED\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["DIV BY ZERO GUARDED"]);
}
