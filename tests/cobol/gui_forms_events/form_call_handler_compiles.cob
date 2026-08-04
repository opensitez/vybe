*> vybe-test: cobol/gui_forms_events/form_call_handler_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC 9 VALUE 1.
PROCEDURE DIVISION.
    CALL "FORM-HANDLER" USING C.
    STOP RUN.

