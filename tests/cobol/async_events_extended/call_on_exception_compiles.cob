*> vybe-test: cobol/async_events_extended/call_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-C.
PROCEDURE DIVISION.
    CALL "SUBC"
        ON EXCEPTION DISPLAY "E"
        NOT ON EXCEPTION DISPLAY "O"
    END-CALL.
    STOP RUN.

