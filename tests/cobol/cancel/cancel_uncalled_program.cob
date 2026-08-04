*> vybe-test: cobol/cancel/cancel_uncalled_program
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CANCEL "never-called-prog"
           DISPLAY "no error expected"
           STOP RUN.

