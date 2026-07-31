use crate::helpers;

#[test]
fn test_loop_nested_varying_after() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOOP-NEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC 9 VALUE 0.
       01 J PIC 9 VALUE 0.
       01 K PIC 9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2
              AFTER J FROM 1 BY 1 UNTIL J > 2
              AFTER K FROM 1 BY 1 UNTIL K > 2
                 ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // 2 * 2 * 2 = 8
    assert_eq!(helpers::run_prints(src), vec!["08"]);
}

#[test]
fn test_loop_varying_fractional() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOOP-FRAC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC 9V9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 0.5 BY 0.5 UNTIL I > 2.0
              ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // 0.5, 1.0, 1.5, 2.0. That's 4 iterations.
    assert_eq!(helpers::run_prints(src), vec!["04"]);
}

#[test]
fn test_loop_varying_negative() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOOP-NEG.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC S9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 5 BY -1 UNTIL I < 1
              ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // 5, 4, 3, 2, 1 = 5 iterations.
    assert_eq!(helpers::run_prints(src), vec!["05"]);
}

#[test]
fn test_loop_exit_perform_cycle() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOOP-CYCLE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC 9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5
              IF I = 3
                 EXIT PERFORM CYCLE
              END-IF
              ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // 1, 2, 4, 5. So 4 times.
    assert_eq!(helpers::run_prints(src), vec!["04"]);
}

#[test]
fn test_loop_exit_perform_early_break() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOOP-EXIT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC 9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 1 BY 1 UNTIL I > 6
              IF I = 4
                 EXIT PERFORM
              END-IF
              ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // executes for I=1,2,3
    assert_eq!(helpers::run_prints(src), vec!["03"]);
}

#[test]
fn test_loop_no_iterations_zero() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOOP-ZERO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC 9 VALUE 0.
       01 TOTAL PIC 99 VALUE 99.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 5 BY 1 UNTIL I < 5
              ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // condition false at start.
    assert_eq!(helpers::run_prints(src), vec!["99"]);
}
