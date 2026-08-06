*> vybe-test: cobol/gui_screen_workflows/screen_section_basic_form_compiles
*> origin: languages/cobol/tests/cobol/test_gui_screen_workflows.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. SCREEN-A.
DATA DIVISION.
SCREEN SECTION.
01 MAIN-SCREEN.
   05 BLANK SCREEN.
   05 LINE 2 COLUMN 10 VALUE "Name:".
   05 LINE 2 COLUMN 20 PIC X(20) USING WS-NAME.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20).
PROCEDURE DIVISION.
    DISPLAY MAIN-SCREEN.
    ACCEPT MAIN-SCREEN.
    STOP RUN.

