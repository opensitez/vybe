*> vybe-test: cobol/arithmetic_control_flow_matrix/condition_name_set_true_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC X VALUE "N".
   88 ACTIVE VALUE "Y".
PROCEDURE DIVISION.
    SET ACTIVE TO TRUE.
    STOP RUN.

