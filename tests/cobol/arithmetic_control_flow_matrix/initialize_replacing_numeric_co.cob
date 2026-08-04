*> vybe-test: cobol/arithmetic_control_flow_matrix/initialize_replacing_numeric_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G.
   05 A PIC 9 VALUE 5.
   05 B PIC X VALUE "Z".
PROCEDURE DIVISION.
    INITIALIZE G REPLACING NUMERIC DATA BY 9.
    STOP RUN.

