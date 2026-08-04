*> vybe-test: cobol/arithmetic_operations_matrix/mul_case_02
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 4.
PROCEDURE DIVISION.
    MULTIPLY 2 BY A.
    STOP RUN.

