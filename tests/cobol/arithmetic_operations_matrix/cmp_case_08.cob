*> vybe-test: cobol/arithmetic_operations_matrix/cmp_case_08
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = 1 + 1 NOT ON SIZE ERROR DISPLAY "O" END-COMPUTE.
    STOP RUN.

