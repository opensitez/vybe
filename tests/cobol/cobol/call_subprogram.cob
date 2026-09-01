*> vybe-test: cobol/cobol/call_subprogram
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CALLER.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(5).
PROCEDURE DIVISION.
    CALL "SUBPROG" USING WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.

