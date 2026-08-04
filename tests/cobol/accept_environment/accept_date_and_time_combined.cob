*> vybe-test: cobol/accept_environment/accept_date_and_time_combined
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-date PIC 9(6).
       01 ws-time PIC 9(8).
       01 ws-day  PIC 9(5).
       PROCEDURE DIVISION.
           ACCEPT ws-date FROM DATE
           ACCEPT ws-time FROM TIME
           ACCEPT ws-day  FROM DAY
           DISPLAY ws-date
           DISPLAY ws-time
           DISPLAY ws-day
           STOP RUN.

