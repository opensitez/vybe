*> vybe-test: cobol/control_flow_calls_matrix/eval_case_04
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE A ALSO B WHEN 1 ALSO 2 DISPLAY "M" WHEN OTHER DISPLAY "N" END-EVALUATE.
    STOP RUN.

