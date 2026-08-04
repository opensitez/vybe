*> vybe-test: cobol/control_flow_calls_matrix/perf_case_06
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 3 BY -1 UNTIL I < 1 DISPLAY I END-PERFORM.
    STOP RUN.

