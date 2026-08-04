*> vybe-test: cobol/arithmetic_operations_matrix/sub_case_02
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 20.
PROCEDURE DIVISION.
    SUBTRACT 5 FROM A.
    STOP RUN.

