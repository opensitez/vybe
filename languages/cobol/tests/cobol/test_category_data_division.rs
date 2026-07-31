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

#[test]
fn test_data_filler_in_group() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-FILL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-REC.
          05 FILLER PIC X(3) VALUE "ABC".
          05 WS-FLD PIC X(3) VALUE "DEF".
       PROCEDURE DIVISION.
           DISPLAY WS-REC.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["ABCDEF"]);
}

#[test]
fn test_data_picture_comp_fields() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-PIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 DECIMAL PIC S9(4)V99 VALUE -12.34.
       PROCEDURE DIVISION.
           DISPLAY DECIMAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["-1234"]);
}

#[test]
fn test_data_66_set_condition_name() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-66.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-NUM PIC 9 VALUE 1.
       88 LOW VALUE 0 THRU 1.
       PROCEDURE DIVISION.
           IF LOW
               DISPLAY "LOW"
           END-IF.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["LOW"]);
}

#[test]
fn test_data_value_clause_spaces() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-VAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-A PIC X(3) VALUE SPACE.
       PROCEDURE DIVISION.
           DISPLAY "[" WS-A "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[   ]"]);
}

#[test]
fn test_data_redefines_impact_group_move() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-MOVE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 GROUP.
          05 A PIC X(2) VALUE "HI".
          05 B PIC X(2) VALUE "JO".
       01 NUM REDEFINES GROUP PIC X(4).
       PROCEDURE DIVISION.
           MOVE GROUP TO NUM.
           DISPLAY NUM.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["HIJO"]);
}

#[test]
fn test_data_usage_display_with_pic_comp5() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-COMP5.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC S9(4) COMP-5 VALUE -12.
       PROCEDURE DIVISION.
           DISPLAY VAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["-12"]);
}
