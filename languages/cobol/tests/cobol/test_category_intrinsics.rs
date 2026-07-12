use crate::helpers;

#[test]
fn test_intrinsic_upper_case() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. UPPER-CASE-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "hello".
       01 RES PIC X(10).
       PROCEDURE DIVISION.
           MOVE FUNCTION UPPER-CASE(STR) TO RES.
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["HELLO     "]);
}

#[test]
fn test_intrinsic_length() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LENGTH-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(15) VALUE "TEST".
       01 LEN PIC 9(2).
       PROCEDURE DIVISION.
           COMPUTE LEN = FUNCTION LENGTH(STR).
           DISPLAY LEN.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["15"]);
}

#[test]
fn test_intrinsic_reverse() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. REVERSE-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(5) VALUE "COBOL".
       01 RES PIC X(5).
       PROCEDURE DIVISION.
           MOVE FUNCTION REVERSE(STR) TO RES.
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["LOBOC"]);
}

#[test]
fn test_intrinsic_mean() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MEAN-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 VALS OCCURS 4 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 10 TO VALS(1).
           MOVE 20 TO VALS(2).
           MOVE 30 TO VALS(3).
           MOVE 40 TO VALS(4).
           COMPUTE RES = FUNCTION MEAN(ALL VALS).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["25"]);
}

#[test]
fn test_intrinsic_median() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MEDIAN-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 VALS OCCURS 5 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 50 TO VALS(1).
           MOVE 10 TO VALS(2).
           MOVE 30 TO VALS(3).
           MOVE 20 TO VALS(4).
           MOVE 40 TO VALS(5).
           COMPUTE RES = FUNCTION MEDIAN(ALL VALS).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["30"]);
}

#[test]
fn test_intrinsic_range() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. RANGE-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 VALS OCCURS 5 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 15 TO VALS(1).
           MOVE 05 TO VALS(2).
           MOVE 45 TO VALS(3).
           MOVE 25 TO VALS(4).
           MOVE 35 TO VALS(5).
           COMPUTE RES = FUNCTION RANGE(ALL VALS).
           DISPLAY RES.
           STOP RUN.
    "#;
    // Max is 45, min is 5. Range is 45 - 5 = 40.
    assert_eq!(helpers::run_prints(src), vec!["40"]);
}

#[test]
fn test_intrinsic_numval() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. NUMVAL-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(10) VALUE "  -123.45 ".
       01 RES PIC S9(4)V99.
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION NUMVAL(STR).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["-012345"]);
}

#[test]
fn test_intrinsic_numval_c() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. NUMVAL-C-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STR PIC X(15) VALUE "  $1,234.56CR".
       01 RES PIC S9(4)V99.
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION NUMVAL-C(STR).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["-123456"]);
}

#[test]
fn test_intrinsic_sum() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SUM-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 VALS OCCURS 3 TIMES PIC 99.
       01 RES PIC 999.
       PROCEDURE DIVISION.
           MOVE 10 TO VALS(1).
           MOVE 20 TO VALS(2).
           MOVE 30 TO VALS(3).
           COMPUTE RES = FUNCTION SUM(ALL VALS).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["060"]);
}

#[test]
fn test_intrinsic_max() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MAX-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 VALS OCCURS 4 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 15 TO VALS(1).
           MOVE 85 TO VALS(2).
           MOVE 45 TO VALS(3).
           MOVE 25 TO VALS(4).
           COMPUTE RES = FUNCTION MAX(ALL VALS).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["85"]);
}

#[test]
fn test_intrinsic_min() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MIN-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 VALS OCCURS 4 TIMES PIC 99.
       01 RES PIC 99.
       PROCEDURE DIVISION.
           MOVE 15 TO VALS(1).
           MOVE 85 TO VALS(2).
           MOVE 05 TO VALS(3).
           MOVE 25 TO VALS(4).
           COMPUTE RES = FUNCTION MIN(ALL VALS).
           DISPLAY RES.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["05"]);
}

#[test]
fn test_intrinsic_when_compiled() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. COMPILED-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC X(21).
       PROCEDURE DIVISION.
           MOVE FUNCTION WHEN-COMPILED TO RES.
           DISPLAY "COMPILED TRIGGERED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
