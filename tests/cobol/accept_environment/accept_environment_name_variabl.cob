*> vybe-test: cobol/accept_environment/accept_environment_name_variable
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-var-name PIC X(20) VALUE "TMPDIR".
       01 ws-var-val  PIC X(200).
       PROCEDURE DIVISION.
           ACCEPT ws-var-val FROM ENVIRONMENT NAME ws-var-name
           IF ws-var-val = SPACES
               DISPLAY "not set"
           ELSE
               DISPLAY ws-var-val
           END-IF
           STOP RUN.

