use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn go_to_simple_paragraph() {
    let out = run_prints(&prog(
        "",
        r#"    GO TO SHOW-MSG.
    DISPLAY "SKIPPED".
    STOP RUN.
SHOW-MSG.
    DISPLAY "GOTO OK"."#,
    ));
    assert_eq!(out, vec!["GOTO OK"]);
}

fn prog(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn go_to_skips_intervening_display() {
    let out = run_prints(&prog(
        "",
        r#"    GO TO AFTER.
    DISPLAY "BEFORE".
    STOP RUN.
AFTER.
    DISPLAY "AFTER"."#,
    ));
    assert_eq!(out, vec!["AFTER"]);
}

#[test]
fn go_to_depending_on_first_target() {
    let out = run_prints(&prog(
        "01 SEL PIC 9 VALUE 1.",
        r#"    GO TO P1 P2 P3 DEPENDING ON SEL.
    DISPLAY "NONE".
    STOP RUN.
P1.
    DISPLAY "ONE".
    STOP RUN.
P2.
    DISPLAY "TWO".
    STOP RUN.
P3.
    DISPLAY "THREE"."#,
    ));
    assert_eq!(out, vec!["ONE"]);
}

#[test]
fn go_to_depending_on_second_target() {
    let out = run_prints(&prog(
        "01 SEL PIC 9 VALUE 2.",
        r#"    GO TO P1 P2 P3 DEPENDING ON SEL.
    DISPLAY "NONE".
    STOP RUN.
P1.
    DISPLAY "ONE".
    STOP RUN.
P2.
    DISPLAY "TWO".
    STOP RUN.
P3.
    DISPLAY "THREE"."#,
    ));
    assert_eq!(out, vec!["TWO"]);
}

#[test]
fn go_to_depending_on_third_target() {
    let out = run_prints(&prog(
        "01 SEL PIC 9 VALUE 3.",
        r#"    GO TO P1 P2 P3 DEPENDING ON SEL.
    DISPLAY "NONE".
    STOP RUN.
P1.
    DISPLAY "ONE".
    STOP RUN.
P2.
    DISPLAY "TWO".
    STOP RUN.
P3.
    DISPLAY "THREE"."#,
    ));
    assert_eq!(out, vec!["THREE"]);
}

#[test]
fn go_to_depending_on_out_of_range_falls_through() {
    let out = run_prints(&prog(
        "01 SEL PIC 9 VALUE 9.",
        r#"    GO TO P1 P2 P3 DEPENDING ON SEL.
    DISPLAY "FALLTHROUGH".
    STOP RUN.
P1.
    DISPLAY "ONE".
    STOP RUN.
P2.
    DISPLAY "TWO".
    STOP RUN.
P3.
    DISPLAY "THREE"."#,
    ));
    assert_eq!(out, vec!["FALLTHROUGH"]);
}

#[test]
fn go_to_from_conditional_branch() {
    let out = run_prints(&prog(
        "01 N PIC 9 VALUE 5.",
        r#"    IF N > 3
        GO TO BIG-LABEL
    END-IF.
    DISPLAY "SMALL".
    STOP RUN.
BIG-LABEL.
    DISPLAY "BIG"."#,
    ));
    assert_eq!(out, vec!["BIG"]);
}

#[test]
fn go_to_past_multiple_paragraphs() {
    let out = run_prints(&prog(
        "",
        r#"    GO TO END-PARA.
PA.
    DISPLAY "PA".
PB.
    DISPLAY "PB".
END-PARA.
    DISPLAY "END"."#,
    ));
    assert_eq!(out, vec!["END"]);
}

#[test]
fn go_to_depending_on_one_target() {
    let out = run_prints(&prog(
        "01 SEL PIC 9 VALUE 1.",
        r#"    GO TO ONLY-TARGET DEPENDING ON SEL.
    DISPLAY "SKIP".
    STOP RUN.
ONLY-TARGET.
    DISPLAY "TARGET"."#,
    ));
    assert_eq!(out, vec!["TARGET"]);
}

#[test]
fn go_to_in_loop_exits_early() {
    let out = run_prints(&prog(
        "01 I PIC 9 VALUE 0.",
        r#"    PERFORM UNTIL I >= 10
        ADD 1 TO I
        IF I = 5
            GO TO DONE
        END-IF
    END-PERFORM.
DONE.
    DISPLAY I."#,
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn go_to_compiles_to_end_of_program() {
    compile_ok(&prog(
        "",
        r#"    GO TO PROGRAM-END.
    DISPLAY "DEAD".
    STOP RUN.
PROGRAM-END.
    DISPLAY "END"."#,
    ));
}

#[test]
fn go_to_depending_two_targets_selects_first() {
    let out = run_prints(&prog(
        "01 S PIC 9 VALUE 1.",
        r#"    GO TO ALPHA BETA DEPENDING ON S.
    STOP RUN.
ALPHA.
    DISPLAY "A".
    STOP RUN.
BETA.
    DISPLAY "B"."#,
    ));
    assert_eq!(out, vec!["A"]);
}

#[test]
fn go_to_depending_four_targets_fourth() {
    let out = run_prints(&prog(
        "01 S PIC 9 VALUE 4.",
        r#"    GO TO P1 P2 P3 P4 DEPENDING ON S.
    DISPLAY "NONE".
    STOP RUN.
P1.
    DISPLAY "ONE". STOP RUN.
P2.
    DISPLAY "TWO". STOP RUN.
P3.
    DISPLAY "THREE". STOP RUN.
P4.
    DISPLAY "FOUR"."#,
    ));
    assert_eq!(out, vec!["FOUR"]);
}

#[test]
fn go_to_after_initialization() {
    let out = run_prints(&prog(
        "01 N PIC 9(3) VALUE 0.",
        r#"    MOVE 42 TO N.
    GO TO SHOW.
    DISPLAY "SKIP".
    STOP RUN.
SHOW.
    DISPLAY N."#,
    ));
    assert_eq!(out, vec!["042"]);
}

#[test]
fn go_to_not_taken_when_condition_false() {
    let out = run_prints(&prog(
        "01 N PIC 9 VALUE 3.",
        r#"    IF N > 10
        GO TO FAR
    END-IF.
    DISPLAY "STAYED".
    STOP RUN.
FAR.
    DISPLAY "FAR"."#,
    ));
    assert_eq!(out, vec!["STAYED"]);
}

#[test]
fn go_to_depending_on_computes_target() {
    let out = run_prints(&prog(
        "01 BASE PIC 9 VALUE 1.\n01 SEL PIC 9 VALUE 0.",
        r#"    ADD BASE TO SEL.
    GO TO FIRST SECOND DEPENDING ON SEL.
    DISPLAY "MISS".
    STOP RUN.
FIRST.
    DISPLAY "FIRST".
    STOP RUN.
SECOND.
    DISPLAY "SECOND"."#,
    ));
    assert_eq!(out, vec!["FIRST"]);
}

#[test]
fn go_to_label_after_perform() {
    let out = run_prints(&prog(
        "01 C PIC 9 VALUE 0.",
        r#"    PERFORM 3 TIMES
        ADD 1 TO C
    END-PERFORM.
    IF C > 5
        GO TO BIG
    END-IF.
    DISPLAY "NORMAL".
    STOP RUN.
BIG.
    DISPLAY "BIG"."#,
    ));
    assert_eq!(out, vec!["NORMAL"]);
}

#[test]
fn go_to_paragraph_in_section() {
    compile_ok(&prog(
        "",
        r#"    PERFORM WORK-SEC.
    STOP RUN.
WORK-SEC SECTION.
    GO TO INNER.
    DISPLAY "SKIPPED".
INNER.
    DISPLAY "INNER"."#,
    ));
}

#[test]
fn go_to_sequence_two_labels() {
    let out = run_prints(&prog(
        "01 FLAG PIC 9 VALUE 1.",
        r#"    GO TO STEP-A STEP-B DEPENDING ON FLAG.
    DISPLAY "FALLTHROUGH".
    STOP RUN.
STEP-A.
    DISPLAY "STEP-A".
    STOP RUN.
STEP-B.
    DISPLAY "STEP-B"."#,
    ));
    assert_eq!(out, vec!["STEP-A"]);
}

#[test]
fn paragraph_and_go_to_computes_sum_then_exits() {
    let out = run_prints(&prog(
        "01 A PIC 9(3) VALUE 40.\n01 B PIC 9(3) VALUE 60.\n01 S PIC 9(4) VALUE 0.",
        r#"    ADD A B GIVING S.
    IF S >= 100
        GO TO RESULT
    END-IF.
    DISPLAY "LESS".
    STOP RUN.
RESULT.
    DISPLAY S."#,
    ));
    assert_eq!(out, vec!["0100"]);
}

#[test]
fn go_to_depending_on_zero_falls_through() {
    let out = run_prints(&prog(
        "01 SEL PIC 9 VALUE 0.",
        r#"    GO TO PA PB PC DEPENDING ON SEL.
    DISPLAY "FALLTHROUGH".
    STOP RUN.
PA.
    DISPLAY "A". STOP RUN.
PB.
    DISPLAY "B". STOP RUN.
PC.
    DISPLAY "C"."#,
    ));
    assert_eq!(out, vec!["FALLTHROUGH"]);
}
