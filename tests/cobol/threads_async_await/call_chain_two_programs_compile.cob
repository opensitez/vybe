*> vybe-test: cobol/threads_async_await/call_chain_two_programs_compiles
*> origin: languages/cobol/tests/cobol/test_threads_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "SUBM".
    CALL "SUBN".
    STOP RUN.

