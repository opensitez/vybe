*> vybe-test: cobol/control_flow_calls_matrix/perf_case_04
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM 3 TIMES ADD 1 TO I END-PERFORM.
    STOP RUN.

