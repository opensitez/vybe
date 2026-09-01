*> vybe-test: cobol/gui_forms_events/gui_button_click_handler_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S19.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 EV PIC X(10) VALUE "CLICK".
PROCEDURE DIVISION.
    CALL "UI-BTN-HANDLER" USING EV.
    STOP RUN.

