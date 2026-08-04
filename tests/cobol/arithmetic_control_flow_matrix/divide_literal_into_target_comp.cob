*> vybe-test: cobol/arithmetic_control_flow_matrix/divide_literal_into_target_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(3) VALUE 16.
PROCEDURE DIVISION.
    DIVIDE 2 INTO R.
    STOP RUN.

