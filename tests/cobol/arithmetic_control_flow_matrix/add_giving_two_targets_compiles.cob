*> vybe-test: cobol/arithmetic_control_flow_matrix/add_giving_two_targets_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 7.
01 B PIC 9(3) VALUE 8.
01 R1 PIC 9(3).
01 R2 PIC 9(3).
PROCEDURE DIVISION.
    ADD A B GIVING R1 R2.
    STOP RUN.

