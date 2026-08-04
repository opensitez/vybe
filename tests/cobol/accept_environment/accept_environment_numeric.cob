*> vybe-test: cobol/accept_environment/accept_environment_numeric
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-timeout-str PIC X(10).
       01 ws-timeout-val PIC 9(5) VALUE 30.
       PROCEDURE DIVISION.
           ACCEPT ws-timeout-str FROM ENVIRONMENT "TIMEOUT_SECS"
           IF ws-timeout-str NOT = SPACES
               MOVE FUNCTION NUMVAL(ws-timeout-str) TO ws-timeout-val
           END-IF
           DISPLAY ws-timeout-val
           STOP RUN.

