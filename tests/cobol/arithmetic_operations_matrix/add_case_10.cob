*> vybe-test: cobol/arithmetic_operations_matrix/add_case_10
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G1.
   05 A PIC 9 VALUE 1.
01 G2.
   05 A PIC 9 VALUE 2.
PROCEDURE DIVISION.
    ADD CORRESPONDING G1 TO G2.
    STOP RUN.

