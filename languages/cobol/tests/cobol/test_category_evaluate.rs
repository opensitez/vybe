use crate::helpers;

#[test]
fn test_evaluate_basic_condition() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-BASIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9 VALUE 2.
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN 1 DISPLAY "ONE"
              WHEN 2 DISPLAY "TWO"
              WHEN 3 DISPLAY "THREE"
              WHEN OTHER DISPLAY "OTHER"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["TWO"]);
}

#[test]
fn test_evaluate_multiple_when() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-MULTI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9 VALUE 4.
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN 1 
              WHEN 3 
              WHEN 5 DISPLAY "ODD"
              WHEN 2
              WHEN 4
              WHEN 6 DISPLAY "EVEN"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["EVEN"]);
}

#[test]
fn test_evaluate_thru_numeric() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-THRU-NUM.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 99 VALUE 15.
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN 01 THRU 10 DISPLAY "1-10"
              WHEN 11 THRU 20 DISPLAY "11-20"
              WHEN 21 THRU 30 DISPLAY "21-30"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["11-20"]);
}

#[test]
fn test_evaluate_thru_alpha() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-THRU-ALPHA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC X VALUE "M".
       PROCEDURE DIVISION.
           EVALUATE VAL
              WHEN "A" THRU "H" DISPLAY "GROUP 1"
              WHEN "I" THRU "P" DISPLAY "GROUP 2"
              WHEN "Q" THRU "Z" DISPLAY "GROUP 3"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["GROUP 2"]);
}

#[test]
fn test_evaluate_true_condition() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-TRUE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL-A PIC 9 VALUE 5.
       01 VAL-B PIC 9 VALUE 10.
       PROCEDURE DIVISION.
           EVALUATE TRUE
              WHEN VAL-A > 10 DISPLAY "A>10"
              WHEN VAL-B > 5  DISPLAY "B>5"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["B>5"]);
}

#[test]
fn test_evaluate_false_condition() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-FALSE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 FLAG PIC X VALUE "Y".
       PROCEDURE DIVISION.
           EVALUATE FALSE
              WHEN FLAG = "Y" DISPLAY "FLAG IS NOT Y"
              WHEN FLAG = "N" DISPLAY "FLAG IS NOT N"
           END-EVALUATE.
           STOP RUN.
    "#;
    // First condition is FALSE (FLAG = "Y" is true).
    // Second condition is TRUE (FLAG = "N" is false).
    // Wait, EVALUATE FALSE matches WHEN <condition> if the condition evaluates to FALSE.
    // FLAG = "Y" is TRUE. Does TRUE match FALSE? No.
    // FLAG = "N" is FALSE. Does FALSE match FALSE? Yes.
    assert_eq!(helpers::run_prints(src), vec!["FLAG IS NOT N"]);
}

#[test]
fn test_evaluate_also_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-ALSO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL-1 PIC X VALUE "A".
       01 VAL-2 PIC 9 VALUE 2.
       PROCEDURE DIVISION.
           EVALUATE VAL-1 ALSO VAL-2
              WHEN "A" ALSO 1 DISPLAY "A1"
              WHEN "A" ALSO 2 DISPLAY "A2"
              WHEN "B" ALSO 2 DISPLAY "B2"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["A2"]);
}

#[test]
fn test_evaluate_also_any() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-ANY-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL-1 PIC X VALUE "C".
       01 VAL-2 PIC 9 VALUE 5.
       PROCEDURE DIVISION.
           EVALUATE VAL-1 ALSO VAL-2
              WHEN "A" ALSO ANY DISPLAY "A-ANY"
              WHEN ANY ALSO 5 DISPLAY "ANY-5"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["ANY-5"]);
}

#[test]
fn test_evaluate_condition_name() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-COND-NAME.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STATUS-CODE PIC X VALUE "P".
          88 IS-ACTIVE VALUE "A".
          88 IS-PENDING VALUE "P".
       PROCEDURE DIVISION.
           EVALUATE TRUE
              WHEN IS-ACTIVE DISPLAY "ACTIVE"
              WHEN IS-PENDING DISPLAY "PENDING"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["PENDING"]);
}

#[test]
fn test_evaluate_partial_also() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EVAL-PARTIAL-ALSO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL-1 PIC X VALUE "X".
       01 VAL-2 PIC X VALUE "Y".
       PROCEDURE DIVISION.
           EVALUATE VAL-1 ALSO VAL-2
              WHEN "X" ALSO "Z" THRU "Y" DISPLAY "BAD"
              WHEN "X" ALSO "A" THRU "Z" DISPLAY "GOOD"
           END-EVALUATE.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["GOOD"]);
}
