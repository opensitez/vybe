*> vybe-test: cobol/promises_async_await/call_chain_three_programs_compiles
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "STEP-A".
    CALL "STEP-B".
    CALL "STEP-C".
    STOP RUN.

