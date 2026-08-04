*> vybe-test: cobol/arithmetic_operations_matrix/cmp_case_02
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 999 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = (2 + 3) * 4.
    STOP RUN.

