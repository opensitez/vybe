*> vybe-test: cobol/gui_forms_events/gui_navigation_next_prev_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S22.
PROCEDURE DIVISION.
    CALL "UI-NEXT".
    CALL "UI-PREV".
    STOP RUN.

