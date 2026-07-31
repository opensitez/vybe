use crate::helpers;

#[test]
fn test_screen_section_basic() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SCREEN-BASIC.
       DATA DIVISION.
       SCREEN SECTION.
       01 CLEAR-SCREEN.
          05 BLANK SCREEN.
       01 INPUT-SCREEN.
          05 LINE 10 COLUMN 20 VALUE "ENTER NAME: " HIGHLIGHT.
          05 LINE 10 COLUMN 35 PIC X(20) TO WS-NAME SECURE.
       WORKING-STORAGE SECTION.
       01 WS-NAME PIC X(20).
       PROCEDURE DIVISION.
           DISPLAY CLEAR-SCREEN.
           ACCEPT INPUT-SCREEN.
           DISPLAY "SCREEN SECTION PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_screen_attributes_color() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SCREEN-COLOR.
       DATA DIVISION.
       SCREEN SECTION.
       01 COLOR-SCREEN.
          05 LINE 5 COL 5 VALUE "RED ON BLUE" 
             FOREGROUND-COLOR 4 BACKGROUND-COLOR 1.
          05 LINE 6 COL 5 VALUE "BLINKING" BLINK REVERSE-VIDEO.
       PROCEDURE DIVISION.
           DISPLAY COLOR-SCREEN.
           DISPLAY "COLORS PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_screen_display_accept_vars() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SCREEN-VARS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-VAR PIC X(10) VALUE "VYBE".
       SCREEN SECTION.
       01 VAR-SCREEN.
          05 LINE 1 COL 1 PIC X(10) FROM WS-VAR.
          05 LINE 2 COL 1 PIC X(10) USING WS-VAR.
       PROCEDURE DIVISION.
           DISPLAY VAR-SCREEN.
           DISPLAY "VARS PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_screen_control() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SCREEN-CTRL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ATTR-VAR PIC 9(4).
       SCREEN SECTION.
       01 DYN-SCREEN.
          05 LINE 1 COL 1 VALUE "DYNAMIC" CONTROL ATTR-VAR.
       PROCEDURE DIVISION.
           MOVE 0001 TO ATTR-VAR.
           DISPLAY DYN-SCREEN.
           DISPLAY "CONTROL PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_screen_section_with_paging() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SCREEN-PAGE.
       DATA DIVISION.
       SCREEN SECTION.
       01 PAGE-SCREEN.
          05 LINE 1 COL 1 VALUE "PAGE 1".
          05 LINE PLUS 2 COL 1 VALUE "NEXT LINE".
       PROCEDURE DIVISION.
           DISPLAY PAGE-SCREEN.
           DISPLAY "PAGING PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_screen_section_with_prompt_input() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SCREEN-PROMPT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS-CODE PIC X(8).
       SCREEN SECTION.
       01 PROMPT-SCREEN.
          05 LINE 1 COL 1 PIC X(8) BEFORE ADVANCING 1.
          05 LINE 2 COL 1 PIC X(8) VALUE "CODE:" PROMPT ">" TO WS-CODE.
       PROCEDURE DIVISION.
           ACCEPT PROMPT-SCREEN.
           DISPLAY "PROMPT PARSED".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}

#[test]
fn test_screen_section_multiple_frames() {
    let src = r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SCREEN-FRAME.
       DATA DIVISION.
       SCREEN SECTION.
       01 FRAME-1.
          05 LINE 1 COL 1 VALUE "A".
       01 FRAME-2.
          05 LINE 3 COL 1 VALUE "B".
       PROCEDURE DIVISION.
           DISPLAY FRAME-1.
           DISPLAY FRAME-2.
           DISPLAY "MULTI FRAME".
           STOP RUN.
    "#;
    let out = helpers::run_prints(src);
    assert!(!out.is_empty());
}
