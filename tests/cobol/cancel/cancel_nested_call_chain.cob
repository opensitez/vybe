*> vybe-test: cobol/cancel/cancel_nested_call_chain
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
           CALL "level-1"
           CANCEL "level-1"
           CALL "level-2"
           CANCEL "level-2"
           CALL "level-3"
           CANCEL "level-3"
           STOP RUN.

