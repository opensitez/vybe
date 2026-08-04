*> vybe-test: cobol/intrinsics/func_char
*> origin: languages/cobol/tests/cobol/test_intrinsics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC X(1).
PROCEDURE DIVISION.
    MOVE FUNCTION CHAR(65) TO C.
    STOP RUN.

