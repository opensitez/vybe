*> vybe-test: cobol/arithmetic_control_flow_matrix/subtract_giving_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 10.
01 B PIC 9(3) VALUE 4.
01 R PIC 9(3).
PROCEDURE DIVISION.
    SUBTRACT B FROM A GIVING R.
    STOP RUN.

