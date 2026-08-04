*> vybe-test: cobol/arithmetic_control_flow_matrix/if_sign_positive_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC S9(3) VALUE 3.
PROCEDURE DIVISION.
    IF X IS POSITIVE DISPLAY "P" END-IF.
    STOP RUN.

