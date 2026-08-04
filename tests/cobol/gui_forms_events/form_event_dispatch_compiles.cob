*> vybe-test: cobol/gui_forms_events/form_event_dispatch_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S17.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 EV PIC X(10) VALUE "CLICK".
PROCEDURE DIVISION.
    CALL "UI-DISPATCH" USING EV.
    STOP RUN.

