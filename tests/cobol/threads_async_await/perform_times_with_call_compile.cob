*> vybe-test: cobol/threads_async_await/perform_times_with_call_compiles
*> origin: languages/cobol/tests/cobol/test_threads_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM 2 TIMES
        CALL "SUBD"
    END-PERFORM.
    STOP RUN.

