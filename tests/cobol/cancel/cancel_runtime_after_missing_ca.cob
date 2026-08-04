*> vybe-test: cobol/cancel/cancel_runtime_after_missing_call_exception
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. T.
       PROCEDURE DIVISION.
           CALL "missing-module"
               ON EXCEPTION
                   DISPLAY "EXC"
           END-CALL
           CANCEL "missing-module"
           DISPLAY "AFTER"
           STOP RUN.

