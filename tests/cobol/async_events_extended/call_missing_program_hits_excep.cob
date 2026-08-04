*> vybe-test: cobol/async_events_extended/call_missing_program_hits_exception
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-EX.
PROCEDURE DIVISION.
    CALL "NONEXIST"
        ON EXCEPTION
            DISPLAY "ERROR"
        NOT ON EXCEPTION
            DISPLAY "SHOULD-NOT"
    END-CALL
    STOP RUN.

