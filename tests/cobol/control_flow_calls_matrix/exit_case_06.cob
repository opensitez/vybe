*> vybe-test: cobol/control_flow_calls_matrix/exit_case_06
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    ALTER L1 TO PROCEED TO L2.
L1. DISPLAY "A".
L2. STOP RUN.

