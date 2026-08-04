*> vybe-test: cobol/arithmetic_control_flow_matrix/call_identifier_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PGM PIC X(8) VALUE "SUBMOD".
PROCEDURE DIVISION.
    CALL PGM.
    STOP RUN.

