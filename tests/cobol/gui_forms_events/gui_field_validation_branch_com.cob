*> vybe-test: cobol/gui_forms_events/gui_field_validation_branch_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S21.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NAME PIC X(20).
PROCEDURE DIVISION.
    ACCEPT NAME.
    IF NAME = SPACES DISPLAY "REQ" ELSE DISPLAY "OK" END-IF.
    STOP RUN.

