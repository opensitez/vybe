*> vybe-test: cobol/gui_forms_events/screen_with_evaluate_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S4.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC 9.
PROCEDURE DIVISION.
    ACCEPT C.
    EVALUATE C
        WHEN 1 DISPLAY "A"
        WHEN 2 DISPLAY "B"
        WHEN OTHER DISPLAY "X"
    END-EVALUATE.
    STOP RUN.

