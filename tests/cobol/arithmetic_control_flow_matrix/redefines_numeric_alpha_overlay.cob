*> vybe-test: cobol/arithmetic_control_flow_matrix/redefines_numeric_alpha_overlay_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4) VALUE 1234.
01 N-X REDEFINES N PIC X(4).
PROCEDURE DIVISION.
    DISPLAY N-X.
    STOP RUN.

