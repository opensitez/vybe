*> vybe-test: cobol/arithmetic_operations_matrix/add_case_05
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 7.
01 B PIC 9 VALUE 1.
PROCEDURE DIVISION.
    ADD B TO A END-ADD.
    STOP RUN.

