*> vybe-test: cobol/accept_environment/accept_environment_path
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-path PIC X(1024).
       PROCEDURE DIVISION.
           ACCEPT ws-path FROM ENVIRONMENT "PATH"
           DISPLAY ws-path
           STOP RUN.

