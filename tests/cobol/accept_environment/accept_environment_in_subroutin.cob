*> vybe-test: cobol/accept_environment/accept_environment_in_subroutine
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. main-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-config PIC X(200).
       PROCEDURE DIVISION.
           CALL "get-config" USING ws-config
           DISPLAY ws-config
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. get-config.
       DATA DIVISION.
       LINKAGE SECTION.
       01 lk-config PIC X(200).
       WORKING-STORAGE SECTION.
       01 ws-val PIC X(200).
       PROCEDURE DIVISION USING lk-config.
           ACCEPT ws-val FROM ENVIRONMENT "APP_CONFIG"
           IF ws-val = SPACES
               MOVE "default.cfg" TO lk-config
           ELSE
               MOVE ws-val TO lk-config
           END-IF
           GOBACK.

