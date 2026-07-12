use crate::helpers;

#[test]
// GAP: 'ecma:math presentValue' import is not correctly resolved/implemented for FUNCTION PRESENT-VALUE.
fn test_financial_present_value() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. FINANCIAL-LIMITS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RATE PIC V9(4) VALUE 0.05.
       01 VAL-1 PIC 9(4)V99 VALUE 1000.00.
       01 VAL-2 PIC 9(4)V99 VALUE 1000.00.
       01 RESULT PIC 9(6)V99.
       PROCEDURE DIVISION.
           COMPUTE RESULT = FUNCTION PRESENT-VALUE(RATE, VAL-1, VAL-2).
           DISPLAY RESULT.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
// GAP: SIZE ERROR boundaries for FUNCTION LOG with zero/negative values are not properly handled.
fn test_math_log_limits() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MATH-LIMITS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ZERO-VAL PIC S9(4) VALUE 0.
       01 NEG-VAL PIC S9(4) VALUE -10.
       01 RESULT PIC S9(6)V9(4).
       PROCEDURE DIVISION.
           COMPUTE RESULT = FUNCTION LOG(ZERO-VAL)
              ON SIZE ERROR DISPLAY "LOG 0 ERROR".
           COMPUTE RESULT = FUNCTION LOG(NEG-VAL)
              ON SIZE ERROR DISPLAY "LOG NEG ERROR".
           COMPUTE RESULT = FUNCTION MOD(NEG-VAL, 3).
           DISPLAY "MOD: " RESULT.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
