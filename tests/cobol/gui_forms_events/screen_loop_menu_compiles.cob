*> vybe-test: cobol/gui_forms_events/screen_loop_menu_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. S3.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL C = 2
        ACCEPT C
        IF C = 1 DISPLAY "ADD" END-IF
    END-PERFORM.
    STOP RUN.

