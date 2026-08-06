*> vybe-test: cobol/gui_screen_workflows/screen_section_menu_compiles
*> origin: languages/cobol/tests/cobol/test_gui_screen_workflows.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. SCREEN-B.
DATA DIVISION.
SCREEN SECTION.
01 MENU-SCREEN.
   05 LINE 1 COLUMN 5 VALUE "1) Add".
   05 LINE 2 COLUMN 5 VALUE "2) Edit".
   05 LINE 3 COLUMN 5 VALUE "3) Exit".
WORKING-STORAGE SECTION.
01 WS-CHOICE PIC 9.
PROCEDURE DIVISION.
    DISPLAY MENU-SCREEN.
    ACCEPT WS-CHOICE.
    STOP RUN.

