*> vybe-test: cobol/accept_environment/accept_environment_with_inspect
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-path     PIC X(500).
       01 ws-colon-ct PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ACCEPT ws-path FROM ENVIRONMENT "PATH"
           INSPECT ws-path
               TALLYING ws-colon-ct FOR ALL ":"
           DISPLAY ws-colon-ct
           STOP RUN.

