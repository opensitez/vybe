*> vybe-test: cobol/arithmetic_operations_matrix/add_case_06
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 9.
01 R1 PIC 99.
01 R2 PIC 99.
PROCEDURE DIVISION.
    ADD A 1 GIVING R1 R2.
    STOP RUN.

