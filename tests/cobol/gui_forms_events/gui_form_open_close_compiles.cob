*> vybe-test: cobol/gui_forms_events/gui_form_open_close_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S20.
PROCEDURE DIVISION.
    CALL "FORM-OPEN".
    CALL "FORM-CLOSE".
    STOP RUN.

