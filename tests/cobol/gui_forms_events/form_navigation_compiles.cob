*> vybe-test: cobol/gui_forms_events/form_navigation_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S16.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PAGE PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF PAGE = 1 DISPLAY "P1" END-IF.
    STOP RUN.

