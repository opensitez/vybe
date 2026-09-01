*> vybe-test: cobol/promises_async_await/call_using_two_arguments_compiles
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
01 A PIC X(10) VALUE "ALICE".
01 B PIC 9(4) VALUE 1001.
PROCEDURE DIVISION.
    CALL "SUB-C" USING A B.
    STOP RUN.

