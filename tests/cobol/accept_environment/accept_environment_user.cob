*> vybe-test: cobol/accept_environment/accept_environment_user
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-user PIC X(64).
       PROCEDURE DIVISION.
           ACCEPT ws-user FROM ENVIRONMENT "USER"
           DISPLAY "User: " ws-user
           STOP RUN.

