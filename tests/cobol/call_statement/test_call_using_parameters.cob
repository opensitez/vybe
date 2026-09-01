*> vybe-test: cobol/call_statement/test_call_using_parameters
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 100.
01 WS-B PIC X(5) VALUE "HELLO".
PROCEDURE DIVISION.

    CALL "SUBPROG" USING BY REFERENCE WS-A
                         BY CONTENT WS-B.
    CALL "SUBPROG" USING BY VALUE WS-A.
    STOP RUN.

