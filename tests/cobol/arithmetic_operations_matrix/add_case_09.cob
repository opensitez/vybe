*> vybe-test: cobol/arithmetic_operations_matrix/add_case_09
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9V9 VALUE 1.5.
01 B PIC 9V9 VALUE 2.5.
01 R PIC 9.
PROCEDURE DIVISION.
    ADD A B GIVING R ROUNDED.
    STOP RUN.

