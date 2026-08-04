*> vybe-test: cobol/arithmetic_control_flow_matrix/compute_nested_expression_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = (3 + 5) * (2 + 1).
    STOP RUN.

