*> vybe-test: cobol/accept_environment/accept_from_day_of_week
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-dow PIC 9.
       PROCEDURE DIVISION.
           ACCEPT ws-dow FROM DAY-OF-WEEK
           EVALUATE ws-dow
               WHEN 1 DISPLAY "Monday"
               WHEN 2 DISPLAY "Tuesday"
               WHEN 3 DISPLAY "Wednesday"
               WHEN 4 DISPLAY "Thursday"
               WHEN 5 DISPLAY "Friday"
               WHEN 6 DISPLAY "Saturday"
               WHEN 7 DISPLAY "Sunday"
           END-EVALUATE
           STOP RUN.

