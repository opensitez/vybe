*> vybe-test: cobol/cobol2023_intrinsics/chained_intrinsics
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(5).
01 WS-VAL PIC 9(5) VALUE 25.
PROCEDURE DIVISION.
    COMPUTE WS-RESULT = FUNCTION ABS(FUNCTION SQRT(WS-VAL)).
    DISPLAY WS-RESULT.
    STOP RUN.

