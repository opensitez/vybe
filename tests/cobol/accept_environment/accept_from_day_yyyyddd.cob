*> vybe-test: cobol/accept_environment/accept_from_day_yyyyddd
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-day PIC 9(7).
       PROCEDURE DIVISION.
           ACCEPT ws-day FROM DAY YYYYDDD
           DISPLAY ws-day
           STOP RUN.

