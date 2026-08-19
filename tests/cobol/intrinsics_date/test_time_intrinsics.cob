*> vybe-test: cobol/intrinsics_date/test_time_intrinsics
*> origin: languages/cobol/tests/cobol/test_intrinsics_date.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SEC PIC 9(9).
PROCEDURE DIVISION.

    COMPUTE WS-SEC = FUNCTION SECONDS-PAST-MIDNIGHT.
    STOP RUN.

