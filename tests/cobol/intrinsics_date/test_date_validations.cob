*> vybe-test: cobol/intrinsics_date/test_date_validations
*> origin: languages/cobol/tests/cobol/test_intrinsics_date.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RES PIC 9(9).
PROCEDURE DIVISION.

    COMPUTE WS-RES = FUNCTION TEST-DATE-YYYYMMDD(20240229).
    COMPUTE WS-RES = FUNCTION TEST-DAY-YYYYDDD(2024001).
    STOP RUN.

