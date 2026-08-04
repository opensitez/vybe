*> vybe-test: cobol/cancel/cancel_after_end_call_compiles
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       PROCEDURE DIVISION.
           CALL "utility-module"
               ON EXCEPTION
                   DISPLAY "MISSING"
               NOT ON EXCEPTION
                   DISPLAY "LOADED"
           END-CALL
           CANCEL "utility-module"
           STOP RUN.

