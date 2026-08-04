*> vybe-test: cobol/control_flow_calls_matrix/eval_case_03
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 7.
PROCEDURE DIVISION.
    EVALUATE X WHEN 1 THRU 5 DISPLAY "L" WHEN 6 THRU 9 DISPLAY "H" WHEN OTHER DISPLAY "O" END-EVALUATE.
    STOP RUN.

