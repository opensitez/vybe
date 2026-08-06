*> vybe-test: cobol/gui_forms_events/form_field_move_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. S5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NAME PIC X(20).
01 OUT PIC X(20).
PROCEDURE DIVISION.
    ACCEPT NAME.
    MOVE NAME TO OUT.
    DISPLAY OUT.
    STOP RUN.

