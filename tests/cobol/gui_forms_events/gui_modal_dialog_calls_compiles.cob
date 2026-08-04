*> vybe-test: cobol/gui_forms_events/gui_modal_dialog_calls_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S23.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 MSG PIC X(30) VALUE "HELLO".
PROCEDURE DIVISION.
    CALL "UI-MODAL" USING MSG.
    STOP RUN.

