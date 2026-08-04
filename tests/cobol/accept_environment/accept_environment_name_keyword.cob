*> vybe-test: cobol/accept_environment/accept_environment_name_keyword
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-env-name PIC X(30) VALUE "HOME".
       01 ws-env-val  PIC X(200).
       PROCEDURE DIVISION.
           ACCEPT ws-env-val FROM ENVIRONMENT NAME ws-env-name
           DISPLAY ws-env-val
           STOP RUN.

