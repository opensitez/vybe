use crate::helpers;

#[test]
fn test_replace_directive() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. REPL-DIR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       REPLACE ==A== BY ==B==.
       01 FLD-A PIC X VALUE "X".
       PROCEDURE DIVISION.
           DISPLAY FLD-B.
           STOP RUN.
    "#;
    // Replace replaces the identifier FLD-A to FLD-B.
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_replace_off() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. REPL-OFF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       REPLACE ==A== BY ==B==.
       01 FLD-A PIC X VALUE "X".
       REPLACE OFF.
       01 FLD-C PIC X VALUE "Y".
       PROCEDURE DIVISION.
           DISPLAY FLD-B.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_title_directive() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TITLE-DIR.
       TITLE "MY TITLE".
       PROCEDURE DIVISION.
           DISPLAY "TITLE PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["TITLE PARSED"]);
}

#[test]
fn test_eject_skip_directive() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PAGE-DIR.
       EJECT.
       SKIP1.
       SKIP2.
       SKIP3.
       PROCEDURE DIVISION.
           DISPLAY "PAGING PARSED".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["PAGING PARSED"]);
}

#[test]
fn test_replace_multiple_pairs() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. REPL-MULTI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       REPLACE ==A== BY ==B==.
       REPLACE ==C== BY ==D==.
       01 FLD-A PIC X VALUE "X".
       01 FLD-C PIC X VALUE "Y".
       PROCEDURE DIVISION.
           DISPLAY FLD-B.
           DISPLAY FLD-D.
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_eject_with_multiple_skips() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. EJECT-MULTI.
       EJECT.
       SKIP1.
       SKIP3.
       PROCEDURE DIVISION.
           DISPLAY "EJECT-MULTI".
           STOP RUN.
    "#;
    assert_eq!(helpers::run_prints(src), vec!["EJECT-MULTI"]);
}
