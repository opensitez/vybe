use super::helpers::{compile_ok, run_prints};

fn prog(data: &str, procs: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, procs
    )
}

#[test]
fn perform_out_of_line_single_paragraph() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM SHOW-MSG.
    STOP RUN.
SHOW-MSG.
    DISPLAY "HELLO FROM PARA"."#,
    ));
    assert_eq!(out, vec!["HELLO FROM PARA"]);
}

#[test]
fn perform_paragraph_twice() {
    let out = run_prints(&prog(
        "01 C PIC 9 VALUE 0.",
        r#"    PERFORM INC.
    PERFORM INC.
    DISPLAY C.
    STOP RUN.
INC.
    ADD 1 TO C."#,
    ));
    assert_eq!(out, vec!["2"]);
}

#[test]
fn perform_thru_two_paragraphs() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM P1 THRU P2.
    STOP RUN.
P1.
    DISPLAY "P1".
P2.
    DISPLAY "P2"."#,
    ));
    assert_eq!(out, vec!["P1", "P2"]);
}

#[test]
fn perform_paragraph_n_times() {
    let out = run_prints(&prog(
        "01 C PIC 9(2) VALUE 0.",
        r#"    PERFORM ADD-ONE 5 TIMES.
    DISPLAY C.
    STOP RUN.
ADD-ONE.
    ADD 1 TO C."#,
    ));
    assert_eq!(out, vec!["05"]);
}

#[test]
fn perform_paragraph_until() {
    let out = run_prints(&prog(
        "01 N PIC 9(2) VALUE 0.",
        r#"    PERFORM BUMP UNTIL N >= 10.
    DISPLAY N.
    STOP RUN.
BUMP.
    ADD 2 TO N."#,
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn perform_paragraph_varying() {
    let out = run_prints(&prog(
        "01 I PIC 9 VALUE 0.\n01 S PIC 9(3) VALUE 0.",
        r#"    PERFORM ACCUM VARYING I FROM 1 BY 1 UNTIL I > 5.
    DISPLAY S.
    STOP RUN.
ACCUM.
    ADD I TO S."#,
    ));
    // 1+2+3+4+5 = 15
    assert_eq!(out, vec!["15"]);
}

#[test]
fn perform_section_via_name() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM MY-SECTION.
    STOP RUN.
MY-SECTION SECTION.
    DISPLAY "IN SECTION"."#,
    ));
    assert_eq!(out, vec!["IN SECTION"]);
}

#[test]
fn perform_paragraph_with_args_via_working_storage() {
    let out = run_prints(&prog(
        "01 X PIC 9(3) VALUE 7.\n01 Y PIC 9(3) VALUE 0.",
        r#"    PERFORM DOUBLE.
    DISPLAY Y.
    STOP RUN.
DOUBLE.
    MULTIPLY X BY 2 GIVING Y."#,
    ));
    assert_eq!(out, vec!["14"]);
}

#[test]
fn perform_thru_three_paragraphs() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM A THRU C.
    STOP RUN.
A.
    DISPLAY "A".
B.
    DISPLAY "B".
C.
    DISPLAY "C"."#,
    ));
    assert_eq!(out, vec!["A", "B", "C"]);
}

#[test]
fn perform_paragraph_nested_call() {
    let out = run_prints(&prog(
        "01 R PIC 9(3) VALUE 0.",
        r#"    PERFORM OUTER.
    DISPLAY R.
    STOP RUN.
OUTER.
    PERFORM INNER.
    ADD 10 TO R.
INNER.
    ADD 1 TO R."#,
    ));
    assert_eq!(out, vec!["11"]);
}

#[test]
fn perform_paragraph_multiple_times_loop() {
    let out = run_prints(&prog(
        "01 C PIC 9(2) VALUE 0.",
        r#"    PERFORM TICK 10 TIMES.
    DISPLAY C.
    STOP RUN.
TICK.
    ADD 1 TO C."#,
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn perform_thru_only_executes_named_range() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM P1 THRU P2.
    STOP RUN.
P1.
    DISPLAY "ONE".
P2.
    DISPLAY "TWO".
P3.
    DISPLAY "THREE"."#,
    ));
    // P3 should not execute
    assert_eq!(out, vec!["ONE", "TWO"]);
}

#[test]
fn perform_paragraph_until_with_not() {
    let out = run_prints(&prog(
        "01 F PIC 9 VALUE 0.",
        r#"    PERFORM FLIP UNTIL NOT F = 0.
    DISPLAY F.
    STOP RUN.
FLIP.
    ADD 1 TO F."#,
    ));
    assert_eq!(out, vec!["1"]);
}

#[test]
fn perform_paragraph_sets_ws_field() {
    let out = run_prints(&prog(
        "01 NAME PIC X(10) VALUE SPACES.",
        r#"    PERFORM SET-NAME.
    DISPLAY NAME.
    STOP RUN.
SET-NAME.
    MOVE "COBOL" TO NAME."#,
    ));
    assert_eq!(out, vec!["COBOL     "]);
}

#[test]
fn perform_twice_different_paragraphs() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM PA.
    PERFORM PB.
    STOP RUN.
PA.
    DISPLAY "FIRST".
PB.
    DISPLAY "SECOND"."#,
    ));
    assert_eq!(out, vec!["FIRST", "SECOND"]);
}

#[test]
fn perform_paragraph_varying_accumulates_even() {
    let out = run_prints(&prog(
        "01 I PIC 9(2) VALUE 0.\n01 S PIC 9(4) VALUE 0.",
        r#"    PERFORM ADD-EVEN VARYING I FROM 2 BY 2 UNTIL I > 10.
    DISPLAY S.
    STOP RUN.
ADD-EVEN.
    ADD I TO S."#,
    ));
    // 2+4+6+8+10 = 30
    assert_eq!(out, vec!["30"]);
}

#[test]
fn perform_out_of_line_display_counter_each_time() {
    let out = run_prints(&prog(
        "01 C PIC 9 VALUE 0.",
        r#"    PERFORM SHOW-C 3 TIMES.
    STOP RUN.
SHOW-C.
    ADD 1 TO C.
    DISPLAY C."#,
    ));
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn perform_section_with_multiple_paras() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM WORK-SEC.
    STOP RUN.
WORK-SEC SECTION.
    DISPLAY "A".
    DISPLAY "B"."#,
    ));
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn perform_paragraph_computes_factorial() {
    let out = run_prints(&prog(
        "01 N PIC 9 VALUE 5.\n01 F PIC 9(5) VALUE 1.\n01 I PIC 9 VALUE 1.",
        r#"    PERFORM FACT-STEP UNTIL I > N.
    DISPLAY F.
    STOP RUN.
FACT-STEP.
    MULTIPLY I BY F.
    ADD 1 TO I."#,
    ));
    assert_eq!(out, vec!["00120"]);
}

#[test]
fn perform_paragraph_exits_correctly() {
    let out = run_prints(&prog(
        "01 FLAG PIC X VALUE \"N\".",
        r#"    PERFORM CHECK-FLAG.
    DISPLAY "AFTER".
    STOP RUN.
CHECK-FLAG.
    MOVE "Y" TO FLAG."#,
    ));
    assert_eq!(out, vec!["AFTER"]);
}

#[test]
fn perform_paragraph_conditionally() {
    let out = run_prints(&prog(
        "01 X PIC 9 VALUE 1.",
        r#"    IF X = 1
        PERFORM HELLO
    END-IF.
    STOP RUN.
HELLO.
    DISPLAY "YES"."#,
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn perform_thru_skips_para_before_start() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM P2 THRU P3.
    STOP RUN.
P1.
    DISPLAY "SKIP".
P2.
    DISPLAY "TWO".
P3.
    DISPLAY "THREE"."#,
    ));
    assert_eq!(out, vec!["TWO", "THREE"]);
}

#[test]
fn perform_paragraph_loop_with_until_and_display_inside() {
    let out = run_prints(&prog(
        "01 I PIC 9 VALUE 0.",
        r#"    PERFORM STEP UNTIL I >= 3.
    STOP RUN.
STEP.
    ADD 1 TO I.
    DISPLAY I."#,
    ));
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn perform_paragraph_compiles_empty_body() {
    compile_ok(&prog(
        "",
        r#"    PERFORM NOOP.
    STOP RUN.
NOOP.
    CONTINUE."#,
    ));
}

#[test]
fn perform_paragraph_varying_with_display_first_and_last() {
    let out = run_prints(&prog(
        "01 I PIC 9 VALUE 0.",
        r#"    PERFORM SHOW VARYING I FROM 1 BY 1 UNTIL I > 4.
    STOP RUN.
SHOW.
    DISPLAY I."#,
    ));
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn perform_paragraph_called_from_loop() {
    let out = run_prints(&prog(
        "01 I PIC 9 VALUE 0.",
        r#"    PERFORM 3 TIMES
        PERFORM TICK
    END-PERFORM.
    DISPLAY I.
    STOP RUN.
TICK.
    ADD 1 TO I."#,
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn perform_paragraph_from_conditional_branch() {
    let out = run_prints(&prog(
        "01 X PIC 9 VALUE 0.",
        r#"    PERFORM SET-X.
    IF X = 99
        DISPLAY "OK"
    ELSE
        DISPLAY "FAIL"
    END-IF.
    STOP RUN.
SET-X.
    MOVE 99 TO X."#,
    ));
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn perform_with_test_before_compiles() {
    compile_ok(&prog(
        "01 N PIC 9 VALUE 0.",
        r#"    PERFORM WITH TEST BEFORE UNTIL N >= 5
        ADD 1 TO N
    END-PERFORM."#,
    ));
}

#[test]
fn perform_paragraph_until_uses_updated_ws() {
    let out = run_prints(&prog(
        "01 TOTAL PIC 9(4) VALUE 0.\n01 STEP  PIC 9(3) VALUE 100.",
        r#"    PERFORM ADD-STEP UNTIL TOTAL >= 500.
    DISPLAY TOTAL.
    STOP RUN.
ADD-STEP.
    ADD STEP TO TOTAL."#,
    ));
    assert_eq!(out, vec!["0500"]);
}
