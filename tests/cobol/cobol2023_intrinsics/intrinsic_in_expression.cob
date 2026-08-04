*> vybe-test: cobol/cobol2023_intrinsics/intrinsic_in_expression
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 20.
01 WS-RESULT PIC 9(5).
PROCEDURE DIVISION.
    COMPUTE WS-RESULT = FUNCTION MAX(WS-A WS-B) + 5.
    DISPLAY WS-RESULT.
    STOP RUN.

