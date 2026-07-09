use crate::helpers;

#[test]
fn test_data_redefines() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-REDEF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 DATA-ITEM PIC X(4) VALUE "1234".
       01 DATA-NUM REDEFINES DATA-ITEM PIC 9(4).
       PROCEDURE DIVISION.
           DISPLAY DATA-NUM.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["1234"]);
}

#[test]
fn test_data_renames() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-RENAME.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 REC.
          05 FLD-A PIC X VALUE "A".
          05 FLD-B PIC X VALUE "B".
          05 FLD-C PIC X VALUE "C".
       66 ALIAS-AC RENAMES FLD-A THRU FLD-C.
       PROCEDURE DIVISION.
           DISPLAY ALIAS-AC.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["ABC"]);
}

#[test]
fn test_data_blank_when_zero() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-BWZ.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9(3) BLANK WHEN ZERO.
       PROCEDURE DIVISION.
           MOVE 0 TO VAL.
           DISPLAY "[" VAL "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[   ]"]);
}

#[test]
fn test_data_justified_right() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-JUST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC X(5) JUSTIFIED RIGHT.
       PROCEDURE DIVISION.
           MOVE "AB" TO VAL.
           DISPLAY "[" VAL "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[   AB]"]);
}

#[test]
fn test_data_sign_leading_separate() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-SIGN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC S9(3) SIGN IS LEADING SEPARATE CHARACTER.
       PROCEDURE DIVISION.
           MOVE -123 TO VAL.
           DISPLAY VAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["-123"]);
}

#[test]
fn test_data_sign_trailing_separate() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-SIGN-TR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC S9(3) SIGN IS TRAILING SEPARATE CHARACTER.
       PROCEDURE DIVISION.
           MOVE -456 TO VAL.
           DISPLAY VAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["456-"]);
}

#[test]
fn test_data_usage_comp() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-COMP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9(4) USAGE COMP VALUE 1000.
       PROCEDURE DIVISION.
           DISPLAY VAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["1000"]);
}

#[test]
fn test_data_value_hex() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-HEX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC X VALUE X"41".
       PROCEDURE DIVISION.
           DISPLAY VAL.
           STOP RUN.
    "#;
    // Hex 41 is 'A' in ASCII.
    assert_eq!(helpers::run_prints(src), vec!["A"]);
}

#[test]
fn test_data_level_88_multiple() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-88.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 STATUS-CODE PIC X.
          88 IS-VALID VALUE "A" "B" "C".
       PROCEDURE DIVISION.
           MOVE "B" TO STATUS-CODE.
           IF IS-VALID
              DISPLAY "VALID"
           END-IF.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["VALID"]);
}
