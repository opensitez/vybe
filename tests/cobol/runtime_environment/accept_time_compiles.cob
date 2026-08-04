*> vybe-test: cobol/runtime_environment/accept_time_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TIME PIC 9(8).
PROCEDURE DIVISION.
    ACCEPT WS-TIME FROM TIME.
    STOP RUN.

