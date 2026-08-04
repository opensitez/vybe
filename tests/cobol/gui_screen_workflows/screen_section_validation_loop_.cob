*> vybe-test: cobol/gui_screen_workflows/screen_section_validation_loop_compiles
*> origin: languages/cobol/tests/cobol/test_gui_screen_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SCREEN-C.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CHOICE PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL WS-CHOICE = 3
        ACCEPT WS-CHOICE
        EVALUATE WS-CHOICE
            WHEN 1 DISPLAY "ADD"
            WHEN 2 DISPLAY "EDIT"
            WHEN 3 DISPLAY "EXIT"
            WHEN OTHER DISPLAY "INVALID"
        END-EVALUATE
    END-PERFORM.
    STOP RUN.

