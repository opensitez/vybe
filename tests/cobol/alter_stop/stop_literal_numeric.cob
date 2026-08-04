*> vybe-test: cobol/alter_stop/stop_literal_numeric
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           DISPLAY "before pause"
           STOP 0
           DISPLAY "after pause"
           STOP RUN.

