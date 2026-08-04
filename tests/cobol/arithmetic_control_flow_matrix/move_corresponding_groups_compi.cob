*> vybe-test: cobol/arithmetic_control_flow_matrix/move_corresponding_groups_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G1.
   05 A PIC 9 VALUE 1.
   05 B PIC X VALUE "X".
01 G2.
   05 A PIC 9 VALUE 0.
   05 B PIC X VALUE " ".
PROCEDURE DIVISION.
    MOVE CORRESPONDING G1 TO G2.
    STOP RUN.

