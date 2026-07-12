use crate::helpers;

#[test]
fn test_perform_out_of_line_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-BASIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 COUNTER PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM PARA-A.
           DISPLAY "END".
           STOP RUN.
       PARA-A.
           DISPLAY "PARA-A".
           EXIT.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["PARA-A", "END"]);
}

#[test]
fn test_perform_times() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-TIMES.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM 5 TIMES
              ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["05"]);
}

#[test]
fn test_perform_until_test_before() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-UNTIL-BEFORE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 IDX PIC 9 VALUE 5.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM WITH TEST BEFORE UNTIL IDX > 5
              ADD 1 TO TOTAL
              ADD 1 TO IDX
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // Condition IDX > 5 is already true, so the body executes 0 times.
    assert_eq!(helpers::run_prints(src), vec!["00"]);
}

#[test]
fn test_perform_until_test_after() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-UNTIL-AFTER.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 IDX PIC 9 VALUE 5.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM WITH TEST AFTER UNTIL IDX > 5
              ADD 1 TO TOTAL
              ADD 1 TO IDX
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // Condition tested after body. Executes once. IDX becomes 6. Then IDX > 5 is true, loop ends.
    assert_eq!(helpers::run_prints(src), vec!["01"]);
}

#[test]
fn test_perform_varying_single() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-VARYING.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC 9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 1 BY 2 UNTIL I > 5
              ADD I TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // I = 1, 3, 5. TOTAL = 1+3+5 = 9.
    assert_eq!(helpers::run_prints(src), vec!["09"]);
}

#[test]
fn test_perform_varying_after() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-VAR-AFTER.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 I PIC 9 VALUE 0.
       01 J PIC 9 VALUE 0.
       01 TOTAL PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2
              AFTER J FROM 1 BY 1 UNTIL J > 3
                 ADD 1 TO TOTAL
           END-PERFORM.
           DISPLAY TOTAL.
           STOP RUN.
    "#;
    // I = 1,2. J = 1,2,3. Total = 2 * 3 = 6.
    assert_eq!(helpers::run_prints(src), vec!["06"]);
}

#[test]
fn test_perform_thru() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-THRU.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           PERFORM PARA-A THRU PARA-C.
           DISPLAY "MAIN END".
           STOP RUN.
       PARA-A.
           DISPLAY "A".
       PARA-B.
           DISPLAY "B".
       PARA-C.
           DISPLAY "C".
    "#;
    assert_eq!(helpers::run_prints(src), vec!["A", "B", "C", "MAIN END"]);
}
