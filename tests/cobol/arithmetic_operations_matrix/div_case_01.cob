*> vybe-test: cobol/arithmetic_operations_matrix/div_case_01
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 20.
01 B PIC 9 VALUE 5.
PROCEDURE DIVISION.
    DIVIDE B INTO A.
    STOP RUN.

