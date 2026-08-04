*> vybe-test: cobol/runtime_environment/accept_command_line_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV4.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARG PIC X(64).
PROCEDURE DIVISION.
    ACCEPT WS-ARG FROM COMMAND-LINE.
    STOP RUN.

