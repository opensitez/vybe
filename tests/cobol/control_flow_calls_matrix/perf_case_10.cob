*> vybe-test: cobol/control_flow_calls_matrix/perf_case_10
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM 2 TIMES IF I = 0 DISPLAY "A" END-IF ADD 1 TO I END-PERFORM.
    STOP RUN.

