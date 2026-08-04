*> vybe-test: cobol/intrinsics_statistical/test_intrinsics_constants
*> origin: languages/cobol/tests/cobol/test_intrinsics_statistical.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VAL PIC 9V999999.
PROCEDURE DIVISION.

    COMPUTE WS-VAL = FUNCTION PI + FUNCTION E.
    STOP RUN.

