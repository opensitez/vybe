*> vybe-test: cobol/accept_environment/accept_environment_and_display
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-app-mode PIC X(10).
       01 ws-log-level PIC X(10).
       PROCEDURE DIVISION.
           ACCEPT ws-app-mode  FROM ENVIRONMENT "APP_MODE"
           ACCEPT ws-log-level FROM ENVIRONMENT "LOG_LEVEL"
           EVALUATE ws-app-mode
               WHEN "production"
                   DISPLAY "Running in production"
               WHEN "staging"
                   DISPLAY "Running in staging"
               WHEN OTHER
                   DISPLAY "Running in dev mode"
           END-EVALUATE
           STOP RUN.

