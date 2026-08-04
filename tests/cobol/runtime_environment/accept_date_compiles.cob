*> vybe-test: cobol/runtime_environment/accept_date_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC 9(8).
PROCEDURE DIVISION.
    ACCEPT WS-DATE FROM DATE YYYYMMDD.
    STOP RUN.

