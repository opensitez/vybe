*> vybe-test: cobol/cancel/cancel_in_loop
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

