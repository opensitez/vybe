*> vybe-test: cobol/arithmetic_control_flow_matrix/divide_into_giving_remainder_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 7.
01 B PIC 9(3) VALUE 3.
01 Q PIC 9(3).
01 M PIC 9(3).
PROCEDURE DIVISION.
    DIVIDE B INTO A GIVING Q REMAINDER M.
    STOP RUN.

