*> vybe-test: cobol/arithmetic_operations_matrix/sub_case_10
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 0.
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    STOP RUN.

