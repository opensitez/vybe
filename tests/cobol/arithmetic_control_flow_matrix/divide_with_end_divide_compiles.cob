*> vybe-test: cobol/arithmetic_control_flow_matrix/divide_with_end_divide_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 12.
01 B PIC 9(3) VALUE 3.
PROCEDURE DIVISION.
    DIVIDE B INTO A END-DIVIDE.
    STOP RUN.

