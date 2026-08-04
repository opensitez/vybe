*> vybe-test: cobol/runtime_environment/accept_day_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DAY PIC 9(7).
PROCEDURE DIVISION.
    ACCEPT WS-DAY FROM DAY YYYYDDD.
    STOP RUN.

