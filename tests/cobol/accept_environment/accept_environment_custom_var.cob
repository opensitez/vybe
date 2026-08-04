*> vybe-test: cobol/accept_environment/accept_environment_custom_var
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-db-host PIC X(100).
       01 ws-db-port PIC X(10).
       PROCEDURE DIVISION.
           ACCEPT ws-db-host FROM ENVIRONMENT "DB_HOST"
           ACCEPT ws-db-port FROM ENVIRONMENT "DB_PORT"
           IF ws-db-host = SPACES
               MOVE "localhost" TO ws-db-host
           END-IF
           IF ws-db-port = SPACES
               MOVE "5432" TO ws-db-port
           END-IF
           DISPLAY ws-db-host
           DISPLAY ws-db-port
           STOP RUN.

