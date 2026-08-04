*> vybe-test: cobol/accept_environment/accept_environment_multiple
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-home  PIC X(200).
       01 ws-user  PIC X(64).
       01 ws-shell PIC X(100).
       PROCEDURE DIVISION.
           ACCEPT ws-home  FROM ENVIRONMENT "HOME"
           ACCEPT ws-user  FROM ENVIRONMENT "USER"
           ACCEPT ws-shell FROM ENVIRONMENT "SHELL"
           DISPLAY "home:  " ws-home
           DISPLAY "user:  " ws-user
           DISPLAY "shell: " ws-shell
           STOP RUN.

