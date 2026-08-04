*> vybe-test: cobol/control_flow_calls_matrix/if_case_06
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(3) VALUE "123".
PROCEDURE DIVISION.
    IF X IS NUMERIC DISPLAY "N" END-IF.
    STOP RUN.

