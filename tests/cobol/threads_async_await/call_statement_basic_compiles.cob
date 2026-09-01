*> vybe-test: cobol/threads_async_await/call_statement_basic_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_threads_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "SUBA".
    STOP RUN.

