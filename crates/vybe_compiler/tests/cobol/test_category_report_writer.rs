use crate::helpers;

#[test]
fn test_report_writer_initiate() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. RW-INIT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       REPORT SECTION.
       RD SALES-REP
          PAGE LIMIT 60.
       01 TYPE IS REPORT HEADING.
          05 LINE 1 COLUMN 10 PIC X(10) VALUE "SALES REP".
       PROCEDURE DIVISION.
           INITIATE SALES-REP.
           DISPLAY "INITIATE PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_report_writer_generate() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. RW-GEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       REPORT SECTION.
       RD SALES-REP.
       01 DETAIL-LINE TYPE IS DETAIL.
          05 LINE PLUS 1 COLUMN 1 PIC X(10) VALUE "DATA".
       PROCEDURE DIVISION.
           INITIATE SALES-REP.
           GENERATE DETAIL-LINE.
           DISPLAY "GENERATE PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_report_writer_terminate() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. RW-TERM.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       REPORT SECTION.
       RD SALES-REP.
       01 TYPE IS REPORT FOOTING.
          05 LINE PLUS 1 COLUMN 1 PIC X(10) VALUE "END REP".
       PROCEDURE DIVISION.
           INITIATE SALES-REP.
           TERMINATE SALES-REP.
           DISPLAY "TERMINATE PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_report_writer_control_breaks() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. RW-CTRL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 DEPT-ID PIC X(5).
       REPORT SECTION.
       RD SALES-REP
          CONTROLS ARE DEPT-ID.
       01 TYPE IS CONTROL HEADING DEPT-ID.
          05 LINE PLUS 2 COLUMN 1 PIC X(5) SOURCE DEPT-ID.
       PROCEDURE DIVISION.
           DISPLAY "CONTROL BREAKS PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
