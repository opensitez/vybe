*> vybe-test: cobol/gui_forms_events/form_submit_with_call_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S14.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PAY PIC X(50).
PROCEDURE DIVISION.
    CALL "FORM-SUBMIT" USING PAY.
    STOP RUN.

