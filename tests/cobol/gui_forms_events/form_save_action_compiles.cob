*> vybe-test: cobol/gui_forms_events/form_save_action_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S8.
PROCEDURE DIVISION.
    CALL "FORM-SAVE".
    STOP RUN.

