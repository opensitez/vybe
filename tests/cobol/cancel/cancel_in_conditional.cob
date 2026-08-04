*> vybe-test: cobol/cancel/cancel_in_conditional
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-flag PIC X VALUE "Y".
       PROCEDURE DIVISION.
           CALL "temp-module"
           IF ws-flag = "Y"
               CANCEL "temp-module"
           END-IF
           DISPLAY ws-flag
           STOP RUN.

