*> vybe-test: cobol/gui_forms_events/gui_button_click_handler_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S19.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 EV PIC X(10) VALUE "CLICK".
PROCEDURE DIVISION.
    CALL "UI-BTN-HANDLER" USING EV.
    STOP RUN.

