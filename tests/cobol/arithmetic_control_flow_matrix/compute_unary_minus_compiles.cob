*> vybe-test: cobol/arithmetic_control_flow_matrix/compute_unary_minus_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC S9(3) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = -5 + 2.
    STOP RUN.

