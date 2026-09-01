*> vybe-test: cobol/cancel/cancel_resets_initial_program
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           CALL "init-counter" USING ws-count
           DISPLAY ws-count
           CANCEL "init-counter"
           MOVE 0 TO ws-count
           CALL "init-counter" USING ws-count
           DISPLAY ws-count
           STOP RUN.

