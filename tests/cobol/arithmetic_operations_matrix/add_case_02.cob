*> vybe-test: cobol/arithmetic_operations_matrix/add_case_02
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 3.
01 B PIC 9 VALUE 4.
PROCEDURE DIVISION.
    ADD 2 TO B.
    STOP RUN.

