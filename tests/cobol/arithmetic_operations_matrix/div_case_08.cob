*> vybe-test: cobol/arithmetic_operations_matrix/div_case_08
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 18.
01 B PIC 9 VALUE 6.
PROCEDURE DIVISION.
    DIVIDE A BY B GIVING A.
    STOP RUN.

