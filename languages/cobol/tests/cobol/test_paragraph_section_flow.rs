use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

// ── Paragraph flow ──────────────────────────────────────────
#[test]
fn paragraph_basic_execution_order() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM P1.
    PERFORM P2.
    STOP RUN.
P1.
    DISPLAY "FIRST".
P2.
    DISPLAY "SECOND"."#,
    ));
    assert_eq!(out, vec!["FIRST", "SECOND"]);
}

fn prog(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn paragraph_exit_at_end() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM MY-PARA.
    DISPLAY "RETURNED".
    STOP RUN.
MY-PARA.
    DISPLAY "IN PARA".
    EXIT."#,
    ));
    assert_eq!(out, vec!["IN PARA", "RETURNED"]);
}

#[test]
fn section_basic_execution() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM MY-SECT.
    DISPLAY "BACK".
    STOP RUN.
MY-SECT SECTION.
    DISPLAY "IN SECTION"."#,
    ));
    assert_eq!(out, vec!["IN SECTION", "BACK"]);
}

#[test]
fn section_with_multiple_paragraphs() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM MY-SECT.
    STOP RUN.
MY-SECT SECTION.
PA.
    DISPLAY "PA".
PB.
    DISPLAY "PB"."#,
    ));
    assert_eq!(out, vec!["PA", "PB"]);
}

#[test]
fn section_exit_terminates_section() {
    compile_ok(&prog(
        "",
        r#"    PERFORM MY-SECT.
    STOP RUN.
MY-SECT SECTION.
    DISPLAY "START".
    EXIT SECTION.
    DISPLAY "UNREACHABLE"."#,
    ));
}

#[test]
fn paragraph_called_conditionally_true() {
    let out = run_prints(&prog(
        "01 N PIC 9 VALUE 5.",
        r#"    IF N > 0
        PERFORM SAY-POS
    END-IF.
    STOP RUN.
SAY-POS.
    DISPLAY "POSITIVE"."#,
    ));
    assert_eq!(out, vec!["POSITIVE"]);
}

#[test]
fn paragraph_called_conditionally_false() {
    let out = run_prints(&prog(
        "01 N PIC 9 VALUE 0.",
        r#"    IF N > 0
        PERFORM SAY-POS
    ELSE
        PERFORM SAY-ZERO
    END-IF.
    STOP RUN.
SAY-POS.
    DISPLAY "POSITIVE".
SAY-ZERO.
    DISPLAY "ZERO"."#,
    ));
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn paragraph_recursively_via_counter() {
    let out = run_prints(&prog(
        "01 C PIC 9 VALUE 0.",
        r#"    PERFORM COUNT-PARA UNTIL C >= 3.
    DISPLAY C.
    STOP RUN.
COUNT-PARA.
    ADD 1 TO C."#,
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn paragraph_updates_ws_each_call() {
    let out = run_prints(&prog(
        "01 TOTAL PIC 9(4) VALUE 0.",
        r#"    PERFORM ADD-TEN 5 TIMES.
    DISPLAY TOTAL.
    STOP RUN.
ADD-TEN.
    ADD 10 TO TOTAL."#,
    ));
    assert_eq!(out, vec!["0050"]);
}

#[test]
fn section_and_para_thru_range() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM PA THRU PC.
    STOP RUN.
PA.
    DISPLAY "A".
PB.
    DISPLAY "B".
PC.
    DISPLAY "C"."#,
    ));
    assert_eq!(out, vec!["A", "B", "C"]);
}

#[test]
fn paragraph_exit_paragraph_compiles() {
    compile_ok(&prog(
        "01 FLAG PIC 9 VALUE 1.",
        r#"    PERFORM DO-STUFF.
    STOP RUN.
DO-STUFF.
    IF FLAG = 0
        EXIT PARAGRAPH
    END-IF.
    DISPLAY "CONTINUING"."#,
    ));
}

#[test]
fn paragraph_from_two_different_callers() {
    let out = run_prints(&prog(
        "01 C PIC 9 VALUE 0.",
        r#"    PERFORM TICK.
    PERFORM TICK.
    PERFORM TICK.
    DISPLAY C.
    STOP RUN.
TICK.
    ADD 1 TO C."#,
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn section_contains_nested_perform() {
    let out = run_prints(&prog(
        "01 S PIC 9(3) VALUE 0.",
        r#"    PERFORM OUTER-SEC.
    DISPLAY S.
    STOP RUN.
OUTER-SEC SECTION.
    PERFORM INNER-PARA.
    PERFORM INNER-PARA.
INNER-PARA.
    ADD 10 TO S."#,
    ));
    assert_eq!(out, vec!["20"]);
}

#[test]
fn paragraph_varying_calls_n_times() {
    let out = run_prints(&prog(
        "01 I PIC 9 VALUE 0.\n01 PRODUCT PIC 9(5) VALUE 1.",
        r#"    PERFORM DOUBLE VARYING I FROM 1 BY 1 UNTIL I > 5.
    DISPLAY PRODUCT.
    STOP RUN.
DOUBLE.
    MULTIPLY 2 BY PRODUCT."#,
    ));
    // 2^5 = 32
    assert_eq!(out, vec!["00032"]);
}

#[test]
fn main_body_continues_after_perform() {
    let out = run_prints(&prog(
        "",
        r#"    DISPLAY "BEFORE".
    PERFORM MIDDLE.
    DISPLAY "AFTER".
    STOP RUN.
MIDDLE.
    DISPLAY "MIDDLE"."#,
    ));
    assert_eq!(out, vec!["BEFORE", "MIDDLE", "AFTER"]);
}

#[test]
fn paragraph_reads_then_writes_ws() {
    let out = run_prints(&prog(
        "01 INPUT-VAL PIC 9(3) VALUE 25.\n01 DOUBLED PIC 9(4) VALUE 0.",
        r#"    PERFORM CALC.
    DISPLAY DOUBLED.
    STOP RUN.
CALC.
    MULTIPLY INPUT-VAL BY 2 GIVING DOUBLED."#,
    ));
    assert_eq!(out, vec!["50"]);
}

#[test]
fn section_exit_from_inner_para() {
    compile_ok(&prog(
        "",
        r#"    PERFORM WORK-SEC.
    STOP RUN.
WORK-SEC SECTION.
START-WORK.
    DISPLAY "START".
    EXIT SECTION.
REST-WORK.
    DISPLAY "UNREACHABLE"."#,
    ));
}

#[test]
fn perform_thru_does_not_fall_past_end_para() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM P1 THRU P2.
    DISPLAY "MAIN".
    STOP RUN.
P1.
    DISPLAY "P1".
P2.
    DISPLAY "P2".
P3.
    DISPLAY "P3"."#,
    ));
    assert_eq!(out, vec!["P1", "P2", "MAIN"]);
}

#[test]
fn paragraph_exit_before_display() {
    compile_ok(&prog(
        "01 COND PIC 9 VALUE 1.",
        r#"    PERFORM GUARDED.
    STOP RUN.
GUARDED.
    IF COND = 1
        EXIT PARAGRAPH
    END-IF.
    DISPLAY "NOT REACHED"."#,
    ));
}

#[test]
fn paragraph_modifies_flag_tested_in_main() {
    let out = run_prints(&prog(
        "01 READY PIC X VALUE \"N\".",
        r#"    PERFORM INIT.
    IF READY = "Y"
        DISPLAY "READY"
    ELSE
        DISPLAY "NOT READY"
    END-IF.
    STOP RUN.
INIT.
    MOVE "Y" TO READY."#,
    ));
    assert_eq!(out, vec!["READY"]);
}

#[test]
fn section_loop_then_return() {
    let out = run_prints(&prog(
        "01 CNT PIC 9(2) VALUE 0.",
        r#"    PERFORM LOOP-SECT.
    DISPLAY CNT.
    STOP RUN.
LOOP-SECT SECTION.
    PERFORM UNTIL CNT >= 5
        ADD 1 TO CNT
    END-PERFORM."#,
    ));
    assert_eq!(out, vec!["05"]);
}

#[test]
fn paragraph_chain_display_order() {
    let out = run_prints(&prog(
        "",
        r#"    PERFORM A.
    STOP RUN.
A.
    DISPLAY "A".
    PERFORM B.
B.
    DISPLAY "B".
    PERFORM C.
C.
    DISPLAY "C"."#,
    ));
    assert_eq!(out, vec!["A", "B", "C"]);
}

#[test]
fn paragraph_with_condition_skips_display() {
    let out = run_prints(&prog(
        "01 X PIC 9 VALUE 0.",
        r#"    PERFORM MAYBE.
    DISPLAY "DONE".
    STOP RUN.
MAYBE.
    IF X > 0
        DISPLAY "INSIDE"
    END-IF."#,
    ));
    assert_eq!(out, vec!["DONE"]);
}

#[test]
fn section_performs_para_in_order() {
    let out = run_prints(&prog(
        "01 OUT PIC X(10) VALUE SPACES.",
        r#"    PERFORM BUILD-SECT.
    DISPLAY OUT.
    STOP RUN.
BUILD-SECT SECTION.
SET-A.
    MOVE "HELLO" TO OUT.
ADD-SPACE.
    MOVE "HELLO " TO OUT.
ADD-WORLD.
    MOVE "HELLO WOR" TO OUT."#,
    ));
    assert_eq!(out, vec!["HELLO WOR "]);
}

#[test]
fn perform_count_paragraphs_called_correctly() {
    let out = run_prints(&prog(
        "01 A-CNT PIC 9 VALUE 0.\n01 B-CNT PIC 9 VALUE 0.",
        r#"    PERFORM A-PARA 3 TIMES.
    PERFORM B-PARA 2 TIMES.
    DISPLAY A-CNT.
    DISPLAY B-CNT.
    STOP RUN.
A-PARA.
    ADD 1 TO A-CNT.
B-PARA.
    ADD 1 TO B-CNT."#,
    ));
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn paragraph_sets_ws_and_returns_to_main() {
    let out = run_prints(&prog(
        "01 RESULT PIC 9(5) VALUE 0.",
        r#"    MOVE 10 TO RESULT.
    PERFORM SQUARE-IT.
    DISPLAY RESULT.
    STOP RUN.
SQUARE-IT.
    MULTIPLY RESULT BY RESULT."#,
    ));
    assert_eq!(out, vec!["00100"]);
}
