*> vybe-test: cobol/cobol2023_intrinsics/day_of_integer
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DAY PIC 9(7).
PROCEDURE DIVISION.
    COMPUTE WS-DAY = FUNCTION DAY-OF-INTEGER(738886).
    DISPLAY WS-DAY.
    STOP RUN.

