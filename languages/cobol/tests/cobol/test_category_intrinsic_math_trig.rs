use crate::helpers;

#[test]
fn test_intrinsic_trig_functions() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-TRIG.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9V9(4) VALUE 0.5.
       01 RES-SIN PIC S9V9(4).
       01 RES-COS PIC S9V9(4).
       01 RES-TAN PIC S9V9(4).
       PROCEDURE DIVISION.
           COMPUTE RES-SIN = FUNCTION SIN(VAL).
           COMPUTE RES-COS = FUNCTION COS(VAL).
           COMPUTE RES-TAN = FUNCTION TAN(VAL).
           DISPLAY "TRIG PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_intrinsic_inverse_trig() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-INV-TRIG.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9V9(4) VALUE 0.5.
       01 RES-ASIN PIC S9V9(4).
       01 RES-ACOS PIC S9V9(4).
       01 RES-ATAN PIC S9V9(4).
       PROCEDURE DIVISION.
           COMPUTE RES-ASIN = FUNCTION ASIN(VAL).
           COMPUTE RES-ACOS = FUNCTION ACOS(VAL).
           COMPUTE RES-ATAN = FUNCTION ATAN(VAL).
           DISPLAY "INV TRIG PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_intrinsic_exp() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-EXP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES-EXP PIC 9(3)V9(4).
       01 RES-EXP10 PIC 9(3)V9(4).
       PROCEDURE DIVISION.
           COMPUTE RES-EXP = FUNCTION EXP(2).
           COMPUTE RES-EXP10 = FUNCTION EXP10(2).
           DISPLAY "EXP PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_intrinsic_mod_rem() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-MOD.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES-MOD PIC 9(2).
       01 RES-REM PIC S9(2).
       PROCEDURE DIVISION.
           COMPUTE RES-MOD = FUNCTION MOD(10, 3).
           COMPUTE RES-REM = FUNCTION REM(-10, 3).
           DISPLAY RES-MOD " " RES-REM.
           STOP RUN.
    "#;
    // MOD(10,3) is 1. REM(-10,3) is -1.
    assert_eq!(helpers::run_prints(src), vec!["01 -01"]);
}

#[test]
fn test_intrinsic_present_value() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INTRINSIC-PV.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9(4)V99.
       01 PAYMENTS.
          05 P-VAL OCCURS 3 TIMES PIC 999 VALUE 100.
       PROCEDURE DIVISION.
           COMPUTE RES = FUNCTION PRESENT-VALUE(0.05, ALL P-VAL).
           DISPLAY "PV PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
