*> vybe-test: cobol/arithmetic_operations_matrix/add_case_04
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 1.
01 R PIC 9.
PROCEDURE DIVISION.
    ADD A B GIVING R.
    STOP RUN.

