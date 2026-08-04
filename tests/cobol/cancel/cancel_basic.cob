*> vybe-test: cobol/cancel/cancel_basic
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CALL "utility-sub"
           CANCEL "utility-sub"
           DISPLAY "cancelled"
           STOP RUN.

