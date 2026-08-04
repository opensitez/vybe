*> vybe-test: cobol/arithmetic_control_flow_matrix/condition_name_set_false_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC X VALUE "Y".
   88 ACTIVE VALUE "Y".
   88 INACTIVE VALUE "N".
PROCEDURE DIVISION.
    SET ACTIVE TO FALSE.
    STOP RUN.

