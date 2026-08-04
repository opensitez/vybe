*> vybe-test: cobol/control_flow_calls_matrix/eval_case_08
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 0.
PROCEDURE DIVISION.
    EVALUATE TRUE WHEN X = 0 DISPLAY "Z" WHEN OTHER DISPLAY "N" END-EVALUATE.
    STOP RUN.

