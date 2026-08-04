*> vybe-test: cobol/arithmetic_control_flow_matrix/add_corresponding_groups_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G1.
   05 A PIC 9(2) VALUE 11.
   05 B PIC 9(2) VALUE 22.
01 G2.
   05 A PIC 9(2) VALUE 1.
   05 B PIC 9(2) VALUE 2.
PROCEDURE DIVISION.
    ADD CORRESPONDING G1 TO G2.
    STOP RUN.

