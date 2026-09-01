*> vybe-test: cobol/gui_forms_events/form_event_dispatch_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S17.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 EV PIC X(10) VALUE "CLICK".
PROCEDURE DIVISION.
    CALL "UI-DISPATCH" USING EV.
    STOP RUN.

