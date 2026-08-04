*> vybe-test: cobol/cobol2023_intrinsics/present_value
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(7)V99.
PROCEDURE DIVISION.
    COMPUTE WS-RESULT = FUNCTION PRESENT-VALUE(0.08 1000).
    DISPLAY WS-RESULT.
    STOP RUN.

