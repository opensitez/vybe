*> vybe-test: cobol/runtime_environment/accept_environment_value_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV10.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VALUE PIC X(64).
PROCEDURE DIVISION.
    ACCEPT WS-VALUE FROM ENVIRONMENT-VALUE.
    STOP RUN.

