*> vybe-test: cobol/runtime_environment/accept_environment_name_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV9.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(64).
PROCEDURE DIVISION.
    DISPLAY "HOME" UPON ENVIRONMENT-NAME.
    ACCEPT WS-NAME FROM ENVIRONMENT-VALUE.
    STOP RUN.

