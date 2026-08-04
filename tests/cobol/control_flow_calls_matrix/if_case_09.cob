*> vybe-test: cobol/control_flow_calls_matrix/if_case_09
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC S9 VALUE 0.
PROCEDURE DIVISION.
    IF X IS ZERO DISPLAY "Z" END-IF.
    STOP RUN.

