*> vybe-test: cobol/intrinsics/func_when_compiled
*> origin: languages/cobol/tests/cobol/test_intrinsics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION WHEN-COMPILED TO D.
    STOP RUN.

