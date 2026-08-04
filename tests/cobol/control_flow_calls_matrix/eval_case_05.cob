*> vybe-test: cobol/control_flow_calls_matrix/eval_case_05
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 9.
PROCEDURE DIVISION.
    EVALUATE A ALSO B WHEN 5 ALSO ANY DISPLAY "H" WHEN OTHER DISPLAY "N" END-EVALUATE.
    STOP RUN.

