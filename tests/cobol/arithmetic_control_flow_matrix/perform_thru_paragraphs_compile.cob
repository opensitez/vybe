*> vybe-test: cobol/arithmetic_control_flow_matrix/perform_thru_paragraphs_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM P1 THRU P2.
    STOP RUN.
P1. DISPLAY "1".
P2. DISPLAY "2".

