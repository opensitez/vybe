*> vybe-test: cobol/cancel/cancel_in_loop
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
       01 ws-i PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING ws-i FROM 1 BY 1 UNTIL ws-i > 3
               CALL "batch-proc" USING ws-i
               CANCEL "batch-proc"
           END-PERFORM
           DISPLAY "loop done"
           STOP RUN.

