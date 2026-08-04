*> vybe-test: cobol/arithmetic_control_flow_matrix/subtract_corresponding_groups_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G1.
   05 A PIC 9(2) VALUE 9.
   05 B PIC 9(2) VALUE 8.
01 G2.
   05 A PIC 9(2) VALUE 4.
   05 B PIC 9(2) VALUE 3.
PROCEDURE DIVISION.
    SUBTRACT CORRESPONDING G2 FROM G1.
    STOP RUN.

