*> vybe-test: cobol/arithmetic_operations_matrix/mul_case_08
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
01 B PIC 9 VALUE 9.
PROCEDURE DIVISION.
    MULTIPLY A BY B.
    STOP RUN.

