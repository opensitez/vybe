*> vybe-test: cobol/accept_environment/accept_environment_basic
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-home PIC X(200).
       PROCEDURE DIVISION.
           ACCEPT ws-home FROM ENVIRONMENT "HOME"
           DISPLAY ws-home
           STOP RUN.

