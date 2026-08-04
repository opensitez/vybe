*> vybe-test: cobol/arithmetic_control_flow_matrix/if_class_numeric_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(5) VALUE "123".
PROCEDURE DIVISION.
    IF X IS NUMERIC DISPLAY "Y" END-IF.
    STOP RUN.

