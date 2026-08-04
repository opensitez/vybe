*> vybe-test: cobol/arithmetic_operations_matrix/mul_case_04
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 3.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    MULTIPLY A BY B END-MULTIPLY.
    STOP RUN.

