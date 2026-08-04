*> vybe-test: cobol/arithmetic_control_flow_matrix/add_with_end_add_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 1.
01 B PIC 9(3) VALUE 2.
PROCEDURE DIVISION.
    ADD A TO B END-ADD.
    STOP RUN.

