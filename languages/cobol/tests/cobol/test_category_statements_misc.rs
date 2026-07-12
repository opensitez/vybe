use crate::helpers;

#[test]
fn test_misc_initialize() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-INIT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 GRP.
          05 FLD-1 PIC X(5) VALUE "HELLO".
          05 FLD-2 PIC 9(3) VALUE 123.
       PROCEDURE DIVISION.
           INITIALIZE GRP.
           DISPLAY "[" FLD-1 "]" FLD-2.
           STOP RUN.
    "#;
    // INITIALIZE sets alphanumeric to spaces, numeric to zeros.
    assert_eq!(helpers::run_prints(src), vec!["[     ]000"]);
}

#[test]
fn test_misc_initialize_replacing() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-INIT-REPL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 GRP.
          05 FLD-1 PIC X(3).
          05 FLD-2 PIC 9(3).
       PROCEDURE DIVISION.
           INITIALIZE GRP REPLACING ALPHANUMERIC BY "A" NUMERIC BY 9.
           DISPLAY FLD-1 " " FLD-2.
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["AAA 999"]);
}

#[test]
fn test_misc_go_to() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-GOTO.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
       PARA-1.
           GO TO PARA-3.
       PARA-2.
           DISPLAY "SKIPPED".
       PARA-3.
           DISPLAY "REACHED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["REACHED"]);
}

#[test]
fn test_misc_go_to_depending_on() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-GOTO-DEP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9 VALUE 2.
       PROCEDURE DIVISION.
       MAIN-PARA.
           GO TO PARA-1 PARA-2 PARA-3 DEPENDING ON VAL.
           DISPLAY "FALLTHROUGH".
           STOP RUN.
       PARA-1.
           DISPLAY "1".
           STOP RUN.
       PARA-2.
           DISPLAY "2".
           STOP RUN.
       PARA-3.
           DISPLAY "3".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["2"]);
}

#[test]
fn test_misc_continue() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-CONT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           IF 1 = 1
              CONTINUE
           ELSE
              DISPLAY "NO"
           END-IF.
           DISPLAY "YES".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["YES"]);
}

#[test]
fn test_misc_accept_date() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-ACC-DATE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-DATE PIC 9(6).
       PROCEDURE DIVISION.
           ACCEPT WS-DATE FROM DATE.
           DISPLAY "DATE ACCEPTED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_misc_accept_time() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-ACC-TIME.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-TIME PIC 9(8).
       PROCEDURE DIVISION.
           ACCEPT WS-TIME FROM TIME.
           DISPLAY "TIME ACCEPTED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_misc_accept_command_line() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-ACC-CMD.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 CMD PIC X(50).
       PROCEDURE DIVISION.
           ACCEPT CMD FROM COMMAND-LINE.
           DISPLAY "CMD ACCEPTED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
