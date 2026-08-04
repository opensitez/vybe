*> vybe-test: cobol/intrinsics_date/test_year_to_yyyy
*> origin: languages/cobol/tests/cobol/test_intrinsics_date.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-YEAR PIC 9(4).
PROCEDURE DIVISION.

    COMPUTE WS-YEAR = FUNCTION YEAR-TO-YYYY(24 50 2000).
    STOP RUN.

