*> vybe-test: cobol/control_flow_calls_matrix/eval_case_06
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 3.
PROCEDURE DIVISION.
    EVALUATE X WHEN 1 DISPLAY "A" WHEN 2 DISPLAY "B" WHEN 3 DISPLAY "C" WHEN OTHER DISPLAY "Z" END-EVALUATE.
    STOP RUN.

