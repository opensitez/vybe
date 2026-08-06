*> vybe-test: cobol/gui_forms_events/screen_basic_display_accept_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. S1.
DATA DIVISION.
SCREEN SECTION.
01 SCR.
   05 LINE 1 COLUMN 1 VALUE "Name".
WORKING-STORAGE SECTION.
01 N PIC X(20).
PROCEDURE DIVISION.
    DISPLAY SCR.
    ACCEPT N.
    STOP RUN.

