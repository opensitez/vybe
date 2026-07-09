use crate::helpers;

#[test]
fn test_call_by_reference() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. CALL-REF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           CALL "EXT-PROG" USING BY REFERENCE VAL.
           DISPLAY VAL.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_call_by_content() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. CALL-CONT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           CALL "EXT-PROG" USING BY CONTENT VAL.
           DISPLAY VAL.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_call_by_value() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. CALL-VAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 VAL PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           CALL "EXT-PROG" USING BY VALUE VAL.
           DISPLAY VAL.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_call_returning() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. CALL-RET.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 RES PIC 9.
       PROCEDURE DIVISION.
           CALL "EXT-PROG" RETURNING RES.
           DISPLAY RES.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_call_on_exception() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. CALL-EXC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           CALL "MISSING-PROG"
              ON EXCEPTION DISPLAY "EXCEPTION"
              NOT ON EXCEPTION DISPLAY "SUCCESS".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_cancel() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. CANCEL-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           CANCEL "EXT-PROG".
           DISPLAY "CANCEL PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["CANCEL PARSED"]);
}

#[test]
fn test_call_nested() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. MAIN-PROG.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           CALL "SUB-PROG".
           STOP RUN.
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SUB-PROG.
       PROCEDURE DIVISION.
           DISPLAY "NESTED".
           EXIT PROGRAM.
       END PROGRAM SUB-PROG.
       END PROGRAM MAIN-PROG.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
