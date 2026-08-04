*> vybe-test: cobol/runtime_environment/accept_from_environment_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PATH PIC X(64).
PROCEDURE DIVISION.
    ACCEPT WS-PATH FROM ENVIRONMENT "PATH".
    STOP RUN.

