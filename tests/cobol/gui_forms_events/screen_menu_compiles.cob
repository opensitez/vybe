*> vybe-test: cobol/gui_forms_events/screen_menu_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S2.
DATA DIVISION.
SCREEN SECTION.
01 SCR.
   05 LINE 1 COLUMN 1 VALUE "1. Add".
   05 LINE 2 COLUMN 1 VALUE "2. Exit".
WORKING-STORAGE SECTION.
01 C PIC 9.
PROCEDURE DIVISION.
    DISPLAY SCR.
    ACCEPT C.
    STOP RUN.

