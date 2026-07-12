use crate::helpers;

#[test]
fn test_declaratives_use_for_debugging() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DECL-DEBUG.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SOURCE-COMPUTER. COMPUTER WITH DEBUGGING MODE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-VAL PIC 9 VALUE 1.
       PROCEDURE DIVISION.
       DECLARATIVES.
       DEBUG-PROC SECTION.
           USE FOR DEBUGGING ON PARA-A.
       DEBUG-PARA.
           DISPLAY "DEBUG TRIGGERED".
       END DECLARATIVES.
       MAIN SECTION.
       PARA-A.
           DISPLAY "PARA A".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_declaratives_global_use() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. DECL-GLOBAL.
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
           USE GLOBAL AFTER STANDARD EXCEPTION PROCEDURE ON TEST-FILE.
       ERR-PARA.
           DISPLAY "GLOBAL EXCEPTION PARSED".
       END DECLARATIVES.
       MAIN SECTION.
           DISPLAY "MAIN".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
