*> vybe-test: cobol/datetime_and_encoding/accept_date_from_system_compiles
*> origin: languages/cobol/tests/cobol/test_datetime_and_encoding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(8).
PROCEDURE DIVISION.
    ACCEPT WS-DATE FROM DATE.
    STOP RUN.

