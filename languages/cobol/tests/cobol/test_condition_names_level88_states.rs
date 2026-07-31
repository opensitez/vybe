use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn condition_name_set_true_compiles() {
    compile_ok(&p(
        "01 F PIC 9.\n   88 ONN VALUE 1.",
        "    SET ONN TO TRUE.",
    ));
}
#[test]
fn condition_name_set_false_compiles() {
    compile_ok(&p(
        "01 F PIC 9.\n   88 OFF VALUE 0.",
        "    SET OFF TO TRUE.",
    ));
}
#[test]
fn condition_name_if_true_compiles() {
    compile_ok(&p(
        "01 F PIC 9 VALUE 1.\n   88 ONN VALUE 1.",
        "    IF ONN DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn condition_name_if_false_compiles() {
    compile_ok(&p(
        "01 F PIC 9 VALUE 0.\n   88 ONN VALUE 1.",
        "    IF NOT ONN DISPLAY \"N\" END-IF.",
    ));
}
#[test]
fn condition_name_evaluate_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 2.",
        "    EVALUATE S WHEN 1 DISPLAY \"A\" WHEN 2 DISPLAY \"B\" WHEN OTHER DISPLAY \"X\" END-EVALUATE.",
    ));
}
#[test]
fn condition_name_multi_values_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.\n   88 ST-B VALUE 2.",
        "    IF ST-A DISPLAY \"A\" END-IF.",
    ));
}
#[test]
fn condition_name_transition_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.\n   88 ST-B VALUE 2.",
        "    SET ST-B TO TRUE.",
    ));
}
#[test]
fn condition_name_display_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    IF ST-A DISPLAY \"OK\" END-IF.",
    ));
}
#[test]
fn condition_name_loop_compiles() {
    compile_ok(&p(
        "01 N PIC 9 VALUE 0.\n01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    PERFORM UNTIL N >= 2\n        ADD 1 TO N\n        IF ST-A DISPLAY \"A\" END-IF\n    END-PERFORM.",
    ));
}
#[test]
fn condition_name_and_if_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 1.\n01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    IF A = 1 AND ST-A DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn condition_name_or_if_compiles() {
    compile_ok(&p(
        "01 A PIC 9 VALUE 0.\n01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    IF A = 1 OR ST-A DISPLAY \"Y\" END-IF.",
    ));
}
#[test]
fn condition_name_not_if_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 0.\n   88 ST-A VALUE 1.",
        "    IF NOT ST-A DISPLAY \"N\" END-IF.",
    ));
}
#[test]
fn condition_name_in_perform_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    PERFORM 2 TIMES IF ST-A DISPLAY \"A\" END-IF END-PERFORM.",
    ));
}
#[test]
fn condition_name_in_evaluate_true_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    EVALUATE TRUE WHEN ST-A DISPLAY \"A\" WHEN OTHER DISPLAY \"X\" END-EVALUATE.",
    ));
}
#[test]
fn condition_name_move_and_set_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 0.\n   88 ST-A VALUE 1.",
        "    MOVE 1 TO S.\n    IF ST-A DISPLAY \"A\" END-IF.",
    ));
}
#[test]
fn condition_name_with_call_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    IF ST-A CALL \"DO-A\" END-IF.",
    ));
}
#[test]
fn condition_name_with_compute_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.\n01 N PIC 9.",
        "    IF ST-A COMPUTE N = 1 + 1 END-IF.",
    ));
}
#[test]
fn condition_name_with_display_chain_compiles() {
    compile_ok(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    IF ST-A DISPLAY \"A\" \"B\" END-IF.",
    ));
}

#[test]
fn condition_name_runtime_true_branch_prints_yes() {
    let output = run_prints(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    IF ST-A DISPLAY \"YES\" ELSE DISPLAY \"NO\" END-IF.",
    ));
    assert_eq!(output, vec!["YES"]);
}

#[test]
fn condition_name_runtime_false_branch_prints_no() {
    let output = run_prints(&p(
        "01 S PIC 9 VALUE 2.\n   88 ST-A VALUE 1.",
        "    IF ST-A DISPLAY \"YES\" ELSE DISPLAY \"NO\" END-IF.",
    ));
    assert_eq!(output, vec!["NO"]);
}

#[test]
fn condition_name_runtime_set_true_updates_storage() {
    let output = run_prints(&p(
        "01 S PIC 9 VALUE 0.\n   88 ST-A VALUE 1.",
        "    SET ST-A TO TRUE.\n    DISPLAY S.",
    ));
    assert_eq!(output, vec!["1"]);
}

#[test]
fn condition_name_runtime_evaluate_true_selects_expected_when() {
    let output = run_prints(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.\n   88 ST-B VALUE 2.",
        "    EVALUATE TRUE WHEN ST-B DISPLAY \"B\" WHEN ST-A DISPLAY \"A\" WHEN OTHER DISPLAY \"X\" END-EVALUATE.",
    ));
    assert_eq!(output, vec!["A"]);
}

#[test]
fn condition_name_runtime_recomputed_after_move() {
    let output = run_prints(&p(
        "01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    MOVE 0 TO S.\n    IF ST-A DISPLAY \"A\" ELSE DISPLAY \"Z\" END-IF.",
    ));
    assert_eq!(output, vec!["Z"]);
}

#[test]
fn condition_name_runtime_boolean_composition_with_and() {
    let output = run_prints(&p(
        "01 FLAG PIC 9 VALUE 1.\n01 S PIC 9 VALUE 1.\n   88 ST-A VALUE 1.",
        "    IF FLAG = 1 AND ST-A DISPLAY \"BOTH\" ELSE DISPLAY \"MISS\" END-IF.",
    ));
    assert_eq!(output, vec!["BOTH"]);
}

#[test]
fn condition_name_with_false_clause_set_false_updates_storage_and_condition() {
    let output = run_prints(&p(
        "01 SW PIC 9 VALUE 1.\n   88 ENABLED VALUE 1 WHEN SET TO FALSE IS 0.",
        "    SET ENABLED TO FALSE.\n    DISPLAY SW.\n    IF ENABLED DISPLAY \"Y\" ELSE DISPLAY \"N\" END-IF.",
    ));
    assert_eq!(output, vec!["0", "N"]);
}

#[test]
fn condition_name_with_multiple_values_true_after_move() {
    let output = run_prints(&p(
        "01 ST PIC 9 VALUE 0.\n   88 OK-STATE VALUE 1 2 3.",
        "    MOVE 2 TO ST.\n    IF OK-STATE DISPLAY \"OK\" ELSE DISPLAY \"BAD\" END-IF.",
    ));
    assert_eq!(output, vec!["OK"]);
}

#[test]
fn condition_name_set_true_uses_condition_value() {
    let output = run_prints(&p(
        "01 ST PIC 9 VALUE 0.\n   88 ACTIVE VALUE 7.",
        "    SET ACTIVE TO TRUE.\n    DISPLAY ST.",
    ));
    assert_eq!(output, vec!["7"]);
}

#[test]
fn condition_name_re_evaluates_after_arithmetic_change() {
    let output = run_prints(&p(
        "01 ST PIC 9 VALUE 1.\n   88 READY VALUE 1.",
        "    ADD 1 TO ST.\n    IF READY DISPLAY \"READY\" ELSE DISPLAY \"NOT-READY\" END-IF.",
    ));
    assert_eq!(output, vec!["NOT-READY"]);
}

#[test]
fn condition_name_range_surface_compiles() {
    compile_ok(&p(
        "01 ST PIC 9(2) VALUE 10.\n   88 LOW VALUE 1 THRU 5.\n   88 MID VALUE 6 THRU 10.\n   88 HIGH VALUE 11 THRU 20.",
        "    SET MID TO TRUE.",
    ));
}

#[test]
fn condition_name_range_runtime_transitions() {
    let output = run_prints(&p(
        "01 ST PIC 9(2) VALUE 0.\n   88 LOW VALUE 1 THRU 5.\n   88 MID VALUE 6 THRU 10.\n   88 HIGH VALUE 11 THRU 20.",
        "    SET LOW TO TRUE\n    IF LOW DISPLAY \"LOW\" ELSE DISPLAY \"NOT-LOW\" END-IF\n    SET HIGH TO TRUE\n    IF HIGH DISPLAY \"HIGH\" ELSE DISPLAY \"NOT-HIGH\" END-IF",
    ));
    assert_eq!(output, vec!["LOW", "HIGH"]);
}
