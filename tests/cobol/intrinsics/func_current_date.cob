*> vybe-test: cobol/intrinsics/func_current_date
*> origin: languages/cobol/tests/cobol/test_intrinsics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO D.
    STOP RUN.

