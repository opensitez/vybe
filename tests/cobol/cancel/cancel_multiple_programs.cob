*> vybe-test: cobol/cancel/cancel_multiple_programs
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CALL "module-a"
           CALL "module-b"
           CALL "module-c"
           CANCEL "module-a"
           CANCEL "module-b"
           CANCEL "module-c"
           DISPLAY "all cancelled"
           STOP RUN.

