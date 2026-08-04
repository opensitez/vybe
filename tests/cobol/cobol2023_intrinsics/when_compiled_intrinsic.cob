*> vybe-test: cobol/cobol2023_intrinsics/when_compiled_intrinsic
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COMPILED PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION WHEN-COMPILED TO WS-COMPILED.
    DISPLAY WS-COMPILED.
    STOP RUN.

