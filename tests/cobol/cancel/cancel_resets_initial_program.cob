*> vybe-test: cobol/cancel/cancel_resets_initial_program
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

