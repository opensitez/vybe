*> vybe-test: cobol/arithmetic_operations_matrix/sub_case_08
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G1.
   05 A PIC 9 VALUE 8.
01 G2.
   05 A PIC 9 VALUE 2.
PROCEDURE DIVISION.
    SUBTRACT CORRESPONDING G2 FROM G1.
    STOP RUN.

