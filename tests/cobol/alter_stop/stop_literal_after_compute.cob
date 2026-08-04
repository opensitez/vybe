*> vybe-test: cobol/alter_stop/stop_literal_after_compute
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result = 12345 * 2
           STOP "Check result"
           DISPLAY ws-result
           STOP RUN.

