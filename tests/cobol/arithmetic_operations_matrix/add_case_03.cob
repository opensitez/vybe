*> vybe-test: cobol/arithmetic_operations_matrix/add_case_03
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
01 C PIC 9 VALUE 3.
PROCEDURE DIVISION.
    ADD A B TO C.
    STOP RUN.

