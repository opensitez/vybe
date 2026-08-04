*> vybe-test: cobol/control_flow_calls_matrix/eval_case_07
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X VALUE "B".
PROCEDURE DIVISION.
    EVALUATE X WHEN "A" DISPLAY "1" WHEN "B" DISPLAY "2" WHEN OTHER DISPLAY "3" END-EVALUATE.
    STOP RUN.

