*> vybe-test: cobol/arithmetic_operations_matrix/cmp_case_09
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 2.
01 B PIC 9 VALUE 3.
01 R PIC 99 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = A * B + 1.
    STOP RUN.

