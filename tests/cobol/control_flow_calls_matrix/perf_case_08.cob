*> vybe-test: cobol/control_flow_calls_matrix/perf_case_08
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM P1 THRU P2.
    STOP RUN.
P1. DISPLAY "1".
P2. DISPLAY "2".

