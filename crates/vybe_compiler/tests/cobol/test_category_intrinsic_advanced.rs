use crate::helpers;

#[test]
fn test_intrinsic_current_date() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-DATE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC X(21).
       PROCEDURE DIVISION.
           MOVE FUNCTION CURRENT-DATE TO RES.
           DISPLAY "DATE TRIGGERED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_intrinsic_date_of_integer() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-DOI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9(8).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION DATE-OF-INTEGER(1).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["16010101"]);
}

#[test]
fn test_intrinsic_annuity() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-ANN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9V9(4).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION ANNUITY(0.05, 3).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["03672"]);
}

#[test]
fn test_intrinsic_random() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-RND.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC V9(4).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION RANDOM(123).
           DISPLAY "RANDOM PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["RANDOM PARSED"]);
}

#[test]
fn test_intrinsic_factorial() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-FACT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9(4).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION FACTORIAL(5).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["0120"]);
}

#[test]
fn test_intrinsic_log() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-LOG.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9V9(4).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION LOG(2.71828).
           DISPLAY RES.
           STOP RUN.
    "#;
    // Log of e is approx 1
    assert_eq!(helpers::run_prints(src), vec!["09999"]);
}

#[test]
fn test_intrinsic_sqrt() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-SQRT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9(2).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION SQRT(144).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["12"]);
}

#[test]
fn test_intrinsic_integer_part() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-INT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC S9(2).
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION INTEGER-PART(-1.5).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["-01"]);
}
