*> vybe-test: cobol/gui_forms_events/gui_event_loop_dispatch_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S24.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 2
        ADD 1 TO I
        CALL "UI-DISPATCH"
    END-PERFORM.
    STOP RUN.

