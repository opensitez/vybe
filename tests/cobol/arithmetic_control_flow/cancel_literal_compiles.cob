*> vybe-test: cobol/arithmetic_control_flow/cancel_literal_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CANCEL "SUBMOD".
    STOP RUN.

