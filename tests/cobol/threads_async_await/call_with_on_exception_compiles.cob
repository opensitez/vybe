*> vybe-test: cobol/threads_async_await/call_with_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_threads_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "SUBC"
        ON EXCEPTION DISPLAY "E"
        NOT ON EXCEPTION DISPLAY "O"
    END-CALL.
    STOP RUN.

