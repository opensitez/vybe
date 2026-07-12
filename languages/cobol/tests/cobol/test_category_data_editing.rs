use crate::helpers;

#[test]
fn test_data_edit_floating_string() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-FLOAT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(4) VALUE 0045.
       01 EDITED PIC $$$,$$$.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[   $45]"]);
}

#[test]
fn test_data_edit_floating_minus() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-FLOAT-MINUS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC S9(3) VALUE -45.
       01 EDITED PIC ---,---.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[   -45]"]);
}

#[test]
fn test_data_edit_floating_plus() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-FLOAT-PLUS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC S9(3) VALUE 45.
       01 EDITED PIC +++,+++.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[   +45]"]);
}

#[test]
fn test_data_edit_insertion_characters() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-INS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(6) VALUE 123456.
       01 EDITED PIC 99/99/99.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["12/34/56"]);
}

#[test]
fn test_data_edit_complex_combined() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-COMPLEX.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC S9(5)V99 VALUE -1234.56.
       01 EDITED PIC $$$,$$9.99DB.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY "[" EDITED "]".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["[  $1,234.56DB]"]);
}

#[test]
fn test_data_edit_zero_insertion() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EDIT-ZERO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 NUM PIC 9(4) VALUE 1234.
       01 EDITED PIC 990099.
       PROCEDURE DIVISION.
           MOVE NUM TO EDITED.
           DISPLAY EDITED.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["120034"]);
}
