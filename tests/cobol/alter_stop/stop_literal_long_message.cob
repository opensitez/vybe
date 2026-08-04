*> vybe-test: cobol/alter_stop/stop_literal_long_message
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           STOP "This is a longer pause message for the operator"
           STOP RUN.

