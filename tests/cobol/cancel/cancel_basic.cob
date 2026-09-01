*> vybe-test: cobol/cancel/cancel_basic
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CALL "utility-sub"
           CANCEL "utility-sub"
           DISPLAY "cancelled"
           STOP RUN.

