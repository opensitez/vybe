*> vybe-test: cobol/new_features/allocate_stmt
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PTR PIC X(10).
PROCEDURE DIVISION.
    ALLOCATE PTR.
    STOP RUN.

