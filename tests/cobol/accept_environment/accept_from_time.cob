*> vybe-test: cobol/accept_environment/accept_from_time
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-time PIC 9(8).
       PROCEDURE DIVISION.
           ACCEPT ws-time FROM TIME
           DISPLAY ws-time
           STOP RUN.

