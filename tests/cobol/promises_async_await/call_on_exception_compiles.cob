*> vybe-test: cobol/promises_async_await/call_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "SUB-E"
        ON EXCEPTION DISPLAY "ERR"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

