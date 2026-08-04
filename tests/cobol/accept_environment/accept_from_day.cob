*> vybe-test: cobol/accept_environment/accept_from_day
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-day PIC 9(5).
       PROCEDURE DIVISION.
           ACCEPT ws-day FROM DAY
           DISPLAY ws-day
           STOP RUN.

