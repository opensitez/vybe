use crate::helpers;

#[test]
fn test_intrinsic_extract_date_time() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-EXTRACT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 D PIC X(21).
       01 RES PIC 9(4).
       PROCEDURE DIVISION.
           MOVE FUNCTION CURRENT-DATE TO D.
           COMPUTE RES = FUNCTION EXTRACT-DATE-TIME(D, "%Y").
           DISPLAY "EXTRACT PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_intrinsic_test_date_time() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-TEST-DT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9.
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION TEST-DATE-TIME("20230101", "%Y%m%d").
           DISPLAY "TEST DT PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_intrinsic_duration() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-DUR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 D1 PIC X(8) VALUE "20230101".
       01 RES PIC X(8).
       PROCEDURE DIVISION.
           MOVE FUNCTION ADD-DURATION(D1, DAYS 5) TO RES.
           DISPLAY "DURATION PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
