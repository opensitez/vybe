*> vybe-test: cobol/promises_async_await/call_chain_three_programs_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "STEP-A".
    CALL "STEP-B".
    CALL "STEP-C".
    STOP RUN.

