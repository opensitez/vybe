*> vybe-test: cobol/promises_async_await/evaluate_with_call_branches_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE K
        WHEN 1 CALL "P1"
        WHEN 2 CALL "P2"
        WHEN OTHER CALL "PX"
    END-EVALUATE.
    STOP RUN.

