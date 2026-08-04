*> vybe-test: cobol/cancel/cancel_preserves_caller_state
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-local PIC 9(5) VALUE 42.
       PROCEDURE DIVISION.
           CALL "sub-prog"
           CANCEL "sub-prog"
           DISPLAY ws-local
           STOP RUN.

