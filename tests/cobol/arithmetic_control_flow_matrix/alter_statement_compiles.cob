*> vybe-test: cobol/arithmetic_control_flow_matrix/alter_statement_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    ALTER A TO PROCEED TO B.
A. DISPLAY "A".
B. STOP RUN.

