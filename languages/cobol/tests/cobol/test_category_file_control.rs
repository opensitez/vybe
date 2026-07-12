use crate::helpers;

#[test]
fn test_file_control_sequential() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-SEQ.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SEQ-FILE ASSIGN TO "seq.dat"
           ORGANIZATION IS SEQUENTIAL
           ACCESS MODE IS SEQUENTIAL
           FILE STATUS IS WS-STAT.
       DATA DIVISION.
       FILE SECTION.
       FD SEQ-FILE.
       01 SEQ-REC PIC X(10).
       WORKING-STORAGE SECTION.
       01 WS-STAT PIC XX.
       PROCEDURE DIVISION.
           DISPLAY "SEQ FILE PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["SEQ FILE PARSED"]);
}

#[test]
fn test_file_control_indexed() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-IDX.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT IDX-FILE ASSIGN TO "idx.dat"
           ORGANIZATION IS INDEXED
           ACCESS MODE IS RANDOM
           RECORD KEY IS KEY-FLD
           ALTERNATE RECORD KEY IS ALT-KEY WITH DUPLICATES.
       DATA DIVISION.
       FILE SECTION.
       FD IDX-FILE.
       01 IDX-REC.
          05 KEY-FLD PIC X(5).
          05 ALT-KEY PIC X(5).
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           DISPLAY "IDX FILE PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["IDX FILE PARSED"]);
}

#[test]
fn test_file_control_relative() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-REL.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT REL-FILE ASSIGN TO "rel.dat"
           ORGANIZATION IS RELATIVE
           ACCESS MODE IS DYNAMIC
           RELATIVE KEY IS REL-KEY.
       DATA DIVISION.
       FILE SECTION.
       FD REL-FILE.
       01 REL-REC PIC X(10).
       WORKING-STORAGE SECTION.
       01 REL-KEY PIC 9(4).
       PROCEDURE DIVISION.
           DISPLAY "REL FILE PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["REL FILE PARSED"]);
}

#[test]
fn test_file_declaratives_error() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-DECL.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT TEST-FILE ASSIGN TO "err.dat"
           FILE STATUS IS WS-STAT.
       DATA DIVISION.
       FILE SECTION.
       FD TEST-FILE.
       01 REC PIC X.
       WORKING-STORAGE SECTION.
       01 WS-STAT PIC XX.
       PROCEDURE DIVISION.
       DECLARATIVES.
       ERR-PROC SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON TEST-FILE.
       ERR-PARA.
           DISPLAY "ERROR CAUGHT " WS-STAT.
       END DECLARATIVES.
       MAIN SECTION.
           DISPLAY "DECLARATIVES PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["DECLARATIVES PARSED"]);
}

#[test]
fn test_file_line_sequential() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-LINE.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT LSEQ-FILE ASSIGN TO "lines.txt"
           ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD LSEQ-FILE.
       01 LREC PIC X(50).
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           DISPLAY "LINE SEQ PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["LINE SEQ PARSED"]);
}
