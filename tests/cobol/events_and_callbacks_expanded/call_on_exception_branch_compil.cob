*> vybe-test: cobol/events_and_callbacks_expanded/call_on_exception_branch_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "SUBX"
        ON EXCEPTION DISPLAY "E"
        NOT ON EXCEPTION DISPLAY "O"
    END-CALL.
    STOP RUN.

