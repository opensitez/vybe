*> vybe-test: cobol/cancel/cancel_after_exception
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-status PIC X VALUE "N".
       PROCEDURE DIVISION.
           CALL "risky-module"
               ON EXCEPTION
                   MOVE "E" TO ws-status
           END-CALL
           CANCEL "risky-module"
           DISPLAY ws-status
           STOP RUN.

