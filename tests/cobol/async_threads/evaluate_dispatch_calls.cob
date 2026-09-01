*> vybe-test: cobol/async_threads/evaluate_dispatch_calls
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC 9 VALUE 1.
PROCEDURE DIVISION.
    EVALUATE K
        WHEN 1 CALL "S1"
        WHEN 2 CALL "S2"
        WHEN OTHER CALL "SX"
    END-EVALUATE.
    STOP RUN.

