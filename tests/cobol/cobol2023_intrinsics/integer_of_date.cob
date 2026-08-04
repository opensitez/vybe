*> vybe-test: cobol/cobol2023_intrinsics/integer_of_date
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC 9(8) VALUE 20240101.
01 WS-INT PIC 9(7).
PROCEDURE DIVISION.
    COMPUTE WS-INT = FUNCTION INTEGER-OF-DATE(WS-DATE).
    DISPLAY WS-INT.
    STOP RUN.

