*> vybe-test: cobol/control_flow_calls_matrix/eval_case_02
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE TRUE WHEN X = 1 DISPLAY "A" WHEN X = 2 DISPLAY "B" WHEN OTHER DISPLAY "Z" END-EVALUATE.
    STOP RUN.

