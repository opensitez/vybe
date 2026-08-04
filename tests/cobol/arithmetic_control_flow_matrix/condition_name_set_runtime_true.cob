*> vybe-test: cobol/arithmetic_control_flow_matrix/condition_name_set_runtime_true_false
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC X VALUE "N".
   88 ACTIVE VALUE "Y".
   88 INACTIVE VALUE "N".
PROCEDURE DIVISION.
    SET ACTIVE TO TRUE
    IF ACTIVE
        DISPLAY "ON"
    END-IF
    SET INACTIVE TO TRUE
    IF NOT ACTIVE
        DISPLAY "OFF"
    END-IF
    STOP RUN.

