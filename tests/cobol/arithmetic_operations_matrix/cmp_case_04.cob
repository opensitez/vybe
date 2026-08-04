*> vybe-test: cobol/arithmetic_operations_matrix/cmp_case_04
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 999 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = 20 / 4.
    STOP RUN.

