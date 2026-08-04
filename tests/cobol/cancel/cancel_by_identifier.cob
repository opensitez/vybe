*> vybe-test: cobol/cancel/cancel_by_identifier
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-prog-name PIC X(20) VALUE "helper-module".
       PROCEDURE DIVISION.
           CALL ws-prog-name
           CANCEL ws-prog-name
           DISPLAY "done"
           STOP RUN.

