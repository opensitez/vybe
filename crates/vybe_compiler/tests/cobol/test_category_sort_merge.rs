use crate::helpers;

#[test]
fn test_sort_using_giving() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SORT-UG.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SORT-WORK ASSIGN TO "work.dat".
           SELECT IN-FILE ASSIGN TO "in.dat".
           SELECT OUT-FILE ASSIGN TO "out.dat".
       DATA DIVISION.
       FILE SECTION.
       SD SORT-WORK.
       01 WORK-REC.
          05 SORT-KEY PIC 9(4).
       FD IN-FILE.
       01 IN-REC PIC X(10).
       FD OUT-FILE.
       01 OUT-REC PIC X(10).
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           SORT SORT-WORK ON ASCENDING KEY SORT-KEY
              USING IN-FILE
              GIVING OUT-FILE.
           DISPLAY "SORT USING GIVING PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["SORT USING GIVING PARSED"]);
}

#[test]
fn test_sort_input_output_procedure() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SORT-PROC.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SORT-WORK ASSIGN TO "work.dat".
       DATA DIVISION.
       FILE SECTION.
       SD SORT-WORK.
       01 WORK-REC.
          05 SORT-KEY PIC 9(4).
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
       MAIN SECTION.
           SORT SORT-WORK ON ASCENDING KEY SORT-KEY
              INPUT PROCEDURE IS IN-PROC
              OUTPUT PROCEDURE IS OUT-PROC.
           DISPLAY "SORT PROCS PARSED".
           STOP RUN.
       IN-PROC SECTION.
           EXIT.
       OUT-PROC SECTION.
           EXIT.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["SORT PROCS PARSED"]);
}

#[test]
fn test_merge_using_giving() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MERGE-UG.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT MERGE-WORK ASSIGN TO "work.dat".
           SELECT IN-1 ASSIGN TO "in1.dat".
           SELECT IN-2 ASSIGN TO "in2.dat".
           SELECT OUT-FILE ASSIGN TO "out.dat".
       DATA DIVISION.
       FILE SECTION.
       SD MERGE-WORK.
       01 WORK-REC.
          05 MERGE-KEY PIC 9(4).
       FD IN-1.
       01 IN1-REC PIC X(10).
       FD IN-2.
       01 IN2-REC PIC X(10).
       FD OUT-FILE.
       01 OUT-REC PIC X(10).
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           MERGE MERGE-WORK ON ASCENDING KEY MERGE-KEY
              USING IN-1 IN-2
              GIVING OUT-FILE.
           DISPLAY "MERGE PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["MERGE PARSED"]);
}

#[test]
fn test_sort_descending_multiple_keys() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SORT-DESC.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SORT-WORK ASSIGN TO "work.dat".
           SELECT IN-FILE ASSIGN TO "in.dat".
           SELECT OUT-FILE ASSIGN TO "out.dat".
       DATA DIVISION.
       FILE SECTION.
       SD SORT-WORK.
       01 WORK-REC.
          05 KEY-1 PIC 9(2).
          05 KEY-2 PIC 9(2).
       FD IN-FILE.
       01 IN-REC PIC X(10).
       FD OUT-FILE.
       01 OUT-REC PIC X(10).
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           SORT SORT-WORK ON ASCENDING KEY KEY-1
                          ON DESCENDING KEY KEY-2
              USING IN-FILE
              GIVING OUT-FILE.
           DISPLAY "SORT DESC PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["SORT DESC PARSED"]);
}
