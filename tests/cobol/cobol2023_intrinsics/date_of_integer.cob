*> vybe-test: cobol/cobol2023_intrinsics/date_of_integer
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC 9(8).
01 WS-INT PIC 9(7) VALUE 738886.
PROCEDURE DIVISION.
    COMPUTE WS-DATE = FUNCTION DATE-OF-INTEGER(WS-INT).
    DISPLAY WS-DATE.
    STOP RUN.

