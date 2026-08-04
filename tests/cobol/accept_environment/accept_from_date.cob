*> vybe-test: cobol/accept_environment/accept_from_date
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-date PIC 9(6).
       PROCEDURE DIVISION.
           ACCEPT ws-date FROM DATE
           DISPLAY ws-date
           STOP RUN.

