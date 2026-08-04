*> vybe-test: cobol/control_flow_calls_matrix/if_case_10
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC S9 VALUE 3.
PROCEDURE DIVISION.
    IF X IS POSITIVE DISPLAY "P" END-IF.
    STOP RUN.

