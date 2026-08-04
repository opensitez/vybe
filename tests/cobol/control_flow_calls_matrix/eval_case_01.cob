*> vybe-test: cobol/control_flow_calls_matrix/eval_case_01
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    EVALUATE X WHEN 1 DISPLAY "A" WHEN OTHER DISPLAY "Z" END-EVALUATE.
    STOP RUN.

