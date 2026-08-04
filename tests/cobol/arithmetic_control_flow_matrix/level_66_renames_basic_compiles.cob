*> vybe-test: cobol/arithmetic_control_flow_matrix/level_66_renames_basic_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G.
   05 A PIC X.
   05 B PIC X.
66 AB RENAMES A THRU B.
PROCEDURE DIVISION.
    DISPLAY AB.
    STOP RUN.

