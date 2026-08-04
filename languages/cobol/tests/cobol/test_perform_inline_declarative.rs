use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn perform_times_zero_body_not_executed() {
    let out = run_prints(&p(
        "",
        "    PERFORM 0 TIMES\n        DISPLAY \"NEVER\"\n    END-PERFORM.\n    DISPLAY \"DONE\".",
    ));
    assert_eq!(out, vec!["DONE"]);
}

#[test]
fn perform_times_once() {
    let out = run_prints(&p(
        "",
        "    PERFORM 1 TIMES\n        DISPLAY \"ONCE\"\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["ONCE"]);
}

#[test]
fn perform_until_test_before_default() {
    // Condition true before body: body never runs
    let out = run_prints(&p(
        "01 K PIC 9 VALUE 5.",
        "    PERFORM UNTIL K >= 5\n        DISPLAY \"BODY\"\n    END-PERFORM.\n    DISPLAY \"AFTER\".",
    ));
    assert_eq!(out, vec!["AFTER"]);
}

#[test]
fn perform_until_with_test_after_runs_once() {
    let out = run_prints(&p(
        "01 K PIC 9 VALUE 5.",
        "    PERFORM WITH TEST AFTER UNTIL K >= 5\n        DISPLAY \"ONCE\"\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["ONCE"]);
}

#[test]
fn perform_varying_counts_to_five() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        DISPLAY I\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["1", "2", "3", "4", "5"]);
}

#[test]
fn perform_varying_by_two() {
    let out = run_prints(&p(
        "01 I PIC 9(2) VALUE 0.",
        "    PERFORM VARYING I FROM 0 BY 2 UNTIL I > 8\n        DISPLAY I\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["00", "02", "04", "06", "08"]);
}

#[test]
fn perform_varying_down_by_one() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 3 BY -1 UNTIL I < 1\n        DISPLAY I\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn perform_until_accumulates_sum() {
    let out = run_prints(&p(
        "01 I PIC 9(2) VALUE 1.\n01 S PIC 9(4) VALUE 0.",
        "    PERFORM UNTIL I > 10\n        ADD I TO S\n        ADD 1 TO I\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["55"]);
}

#[test]
fn perform_with_test_after_increments_before_check() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    PERFORM WITH TEST AFTER UNTIL N >= 3\n        ADD 1 TO N\n    END-PERFORM.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn perform_nested_varying_two_levels() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 J PIC 9 VALUE 0.\n01 S PIC 9(3) VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        PERFORM VARYING J FROM 1 BY 1 UNTIL J > 3\n            ADD 1 TO S\n        END-PERFORM\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["9"]);
}

#[test]
fn perform_times_ten_increments() {
    let out = run_prints(&p(
        "01 C PIC 9(3) VALUE 0.",
        "    PERFORM 10 TIMES\n        ADD 1 TO C\n    END-PERFORM.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn perform_varying_from_zero_inclusive() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 9.",
        "    PERFORM VARYING I FROM 0 BY 1 UNTIL I > 2\n        DISPLAY I\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn perform_until_equality_terminates() {
    let out = run_prints(&p(
        "01 N PIC 9(2) VALUE 0.",
        "    PERFORM UNTIL N = 7\n        ADD 1 TO N\n    END-PERFORM.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["07"]);
}

#[test]
fn perform_varying_preserves_final_index_value() {
    let out = run_prints(&p(
        "01 I PIC 9(2) VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        CONTINUE\n    END-PERFORM.\n    DISPLAY I.",
    ));
    // After loop I = 6 (one past the terminating condition)
    assert_eq!(out, vec!["06"]);
}

#[test]
fn perform_inline_with_display_and_add() {
    let out = run_prints(&p(
        "01 X PIC 9(2) VALUE 0.",
        "    PERFORM 3 TIMES\n        ADD 5 TO X\n    END-PERFORM.\n    DISPLAY X.",
    ));
    assert_eq!(out, vec!["15"]);
}

#[test]
fn perform_varying_with_two_step_and_display_last() {
    let out = run_prints(&p(
        "01 I PIC 9(2) VALUE 0.",
        "    PERFORM VARYING I FROM 2 BY 3 UNTIL I > 11\n        CONTINUE\n    END-PERFORM.\n    DISPLAY I.",
    ));
    // 2, 5, 8, 11 — then 14 > 11 stops; I = 14
    assert_eq!(out, vec!["14"]);
}

#[test]
fn perform_times_variable_count() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 4.\n01 C PIC 9(2) VALUE 0.",
        "    PERFORM N TIMES\n        ADD 1 TO C\n    END-PERFORM.\n    DISPLAY C.",
    ));
    assert_eq!(out, vec!["04"]);
}

#[test]
fn perform_until_with_and_condition() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 0.",
        "    PERFORM UNTIL A >= 3 AND B >= 3\n        ADD 1 TO A\n        ADD 1 TO B\n    END-PERFORM.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn perform_varying_product_accumulation() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.\n01 P PIC 9(4) VALUE 1.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        MULTIPLY I BY P\n    END-PERFORM.\n    DISPLAY P.",
    ));
    // 1*1*2*3*4*5 = 120
    assert_eq!(out, vec!["0120"]);
}

#[test]
fn perform_until_or_exits_on_first_true() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 0.",
        "    PERFORM UNTIL N = 5 OR N = 3\n        ADD 1 TO N\n    END-PERFORM.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn perform_nested_times_loop() {
    let out = run_prints(&p(
        "01 S PIC 9(3) VALUE 0.",
        "    PERFORM 5 TIMES\n        PERFORM 5 TIMES\n            ADD 1 TO S\n        END-PERFORM\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["25"]);
}

#[test]
fn perform_until_changing_step_size() {
    // Double N each iteration: 1,2,4,8 — stop when > 7
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE 1.",
        "    PERFORM UNTIL N > 7\n        MULTIPLY 2 BY N\n    END-PERFORM.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["016"]);
}

#[test]
fn perform_varying_odd_step_boundary() {
    let out = run_prints(&p(
        "01 I PIC 9(2) VALUE 0.\n01 S PIC 9(4) VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 2 UNTIL I > 9\n        ADD I TO S\n    END-PERFORM.\n    DISPLAY S.",
    ));
    // 1+3+5+7+9 = 25
    assert_eq!(out, vec!["25"]);
}

#[test]
fn perform_inline_display_each_iteration() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        DISPLAY I\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn perform_until_not_condition() {
    let out = run_prints(&p(
        "01 F PIC 9 VALUE 0.",
        "    PERFORM UNTIL NOT F = 0\n        ADD 1 TO F\n    END-PERFORM.\n    DISPLAY F.",
    ));
    assert_eq!(out, vec!["1"]);
}

#[test]
fn perform_varying_with_negative_step_halts() {
    let out = run_prints(&p(
        "01 I PIC S9(3) VALUE 0.",
        "    PERFORM VARYING I FROM 10 BY -3 UNTIL I < 1\n        DISPLAY I\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["10", "07", "04", "01"]);
}

#[test]
fn perform_times_displays_index_indirectly() {
    let out = run_prints(&p(
        "01 C PIC 9(2) VALUE 0.",
        "    PERFORM 5 TIMES\n        ADD 2 TO C\n        DISPLAY C\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["02", "04", "06", "08", "10"]);
}

#[test]
fn perform_varying_with_float_step_compiles() {
    compile_ok(&p(
        "01 I PIC 9(3)V9 VALUE 0.",
        "    PERFORM VARYING I FROM 0 BY 0.5 UNTIL I > 2\n        CONTINUE\n    END-PERFORM.",
    ));
}

#[test]
fn perform_until_exit_on_two_conditions_and() {
    // Both counters must move TOWARD the condition, or `AND` never holds: B
    // started at 5 and counted DOWN, so `B > 3` was false from the second pass
    // on and the loop never ended — `cobc -x -free` hangs on the old form too,
    // so it was not a Vybe bug but an invalid program. Counting both up ends at
    // A = B = 4, which is the value this test already asserted.
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 0.",
        "    PERFORM UNTIL A > 3 AND B > 3\n        ADD 1 TO A\n        ADD 1 TO B\n    END-PERFORM.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["4"]);
}

#[test]
fn perform_varying_fills_table() {
    let out = run_prints(&p(
        "01 T.\n   05 E PIC 9 OCCURS 5 TIMES.\n01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        MOVE I TO E(I)\n    END-PERFORM.\n    DISPLAY E(3).",
    ));
    assert_eq!(out, vec!["3"]);
}
