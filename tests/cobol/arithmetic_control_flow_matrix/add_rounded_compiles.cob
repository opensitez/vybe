*> vybe-test: cobol/arithmetic_control_flow_matrix/add_rounded_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9V9 VALUE 1.5.
01 B PIC 9V9 VALUE 2.4.
01 R PIC 9 VALUE 0.
PROCEDURE DIVISION.
    ADD A B GIVING R ROUNDED.
    STOP RUN.

