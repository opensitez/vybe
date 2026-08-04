*> vybe-test: cobol/cobol2023_intrinsics/formatted_date
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION FORMATTED-DATE("YYYY-MM-DD" 20240101) TO WS-DATE.
    DISPLAY WS-DATE.
    STOP RUN.

