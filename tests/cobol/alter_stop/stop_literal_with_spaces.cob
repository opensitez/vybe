*> vybe-test: cobol/alter_stop/stop_literal_with_spaces
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           DISPLAY "starting"
           STOP "   "
           DISPLAY "done"
           STOP RUN.

