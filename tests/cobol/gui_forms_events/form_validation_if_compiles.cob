*> vybe-test: cobol/gui_forms_events/form_validation_if_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 AGE PIC 9(3).
PROCEDURE DIVISION.
    ACCEPT AGE.
    IF AGE < 18 DISPLAY "MINOR" ELSE DISPLAY "ADULT" END-IF.
    STOP RUN.

