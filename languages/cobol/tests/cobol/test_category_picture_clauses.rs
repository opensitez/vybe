use crate::helpers;

#[test]
fn test_pic_alpha() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-A.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC A(5).
       PROCEDURE DIVISION.
           MOVE "ABCDE" TO VAL.
           DISPLAY VAL.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["ABCDE"]);
}

#[test]
fn test_pic_numeric_implied_decimal() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-V.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9(3)V99 VALUE 123.45.
       PROCEDURE DIVISION.
           DISPLAY VAL.
           STOP RUN.
    "#;
    // Without decimal point editing, it displays digits directly: 12345
    assert_eq!(helpers::run_prints(src), vec!["12345"]);
}

#[test]
fn test_pic_editing_decimal_point() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-DOT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(3)V99 VALUE 123.45.
       01 EDITED PIC 999.99.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["123.45"]);
}

#[test]
fn test_pic_editing_comma() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-COMMA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(4) VALUE 1234.
       01 EDITED PIC 9,999.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["1,234"]);
}

#[test]
fn test_pic_editing_zero_suppression() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-Z.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(4) VALUE 0045.
       01 EDITED PIC ZZZ9.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[  45]"]);
}

#[test]
fn test_pic_editing_asterisk() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-STAR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(4) VALUE 0045.
       01 EDITED PIC ****.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["**45"]);
}

#[test]
fn test_pic_editing_currency() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-CURR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(4) VALUE 0123.
       01 EDITED PIC $$$,$$9.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[   $123]"]);
}

#[test]
fn test_pic_editing_plus_minus() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-SIGN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM1 PIC S9(3) VALUE 123.
       01 NUM2 PIC S9(3) VALUE -123.
       01 EDITED1 PIC +999.
       01 EDITED2 PIC -999.
       PROCEDURE DIVISION.
           MOVE NUM1 TO EDITED1.
           MOVE NUM2 TO EDITED2.
           DISPLAY EDITED1 " " EDITED2.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["+123 -123"]);
}

#[test]
fn test_pic_editing_cr_db() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-CRDB.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM1 PIC S9(3) VALUE -123.
       01 NUM2 PIC S9(3) VALUE 123.
       01 EDITED1 PIC 999CR.
       01 EDITED2 PIC 999DB.
       PROCEDURE DIVISION.
           MOVE NUM1 TO EDITED1.
           MOVE NUM2 TO EDITED2.
           DISPLAY EDITED1 " " EDITED2.
           STOP RUN.
    "#;
    // CR and DB only show if negative, otherwise spaces.
    assert_eq!(helpers::run_prints(src), vec!["123CR 123  "]);
}

#[test]
fn test_pic_editing_slash_b_zero() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-INS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(6) VALUE 123456.
       01 EDITED PIC 99/99/99.
       01 NUM2 PIC 9(2) VALUE 12.
       01 EDITED2 PIC 9B9.
       01 EDITED3 PIC 909.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           MOVE NUM2 TO EDITED2.
           MOVE NUM2 TO EDITED3.
           DISPLAY EDITED " " EDITED2 " " EDITED3.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["12/34/56 1 2 102"]);
}

#[test]
fn test_pic_numeric_zero_fill() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-ZFILL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9999 VALUE 12.
       PROCEDURE DIVISION.
           MOVE NUM TO NUM.
           DISPLAY NUM.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["0012"]);
}

#[test]
fn test_pic_comp_2_sign() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-COMP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC S9(3) COMP-2 VALUE -123.
       01 NUM2 PIC S9(3) COMP-2.
       PROCEDURE DIVISION.
           MOVE NUM TO NUM2.
           DISPLAY NUM2.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["-123"]);
}

#[test]
fn test_pic_alpha_num_merge() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PIC-MIX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 A PIC A(3) VALUE "A1 ".
       01 B PIC X(5) VALUE SPACES.
       PROCEDURE DIVISION.
           MOVE A TO B.
           DISPLAY B.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["A1   "]);
}
