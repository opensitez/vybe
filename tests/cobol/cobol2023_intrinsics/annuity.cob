*> vybe-test: cobol/cobol2023_intrinsics/annuity
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(5)V99.
PROCEDURE DIVISION.
    COMPUTE WS-RESULT = FUNCTION ANNUITY(0.05 12).
    DISPLAY WS-RESULT.
    STOP RUN.

