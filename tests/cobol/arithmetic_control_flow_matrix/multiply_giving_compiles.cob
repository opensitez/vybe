*> vybe-test: cobol/arithmetic_control_flow_matrix/multiply_giving_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 6.
01 B PIC 9(3) VALUE 7.
01 R PIC 9(3).
PROCEDURE DIVISION.
    MULTIPLY A BY B GIVING R.
    STOP RUN.

