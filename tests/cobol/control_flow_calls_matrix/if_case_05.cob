*> vybe-test: cobol/control_flow_calls_matrix/if_case_05
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
PROCEDURE DIVISION.
    IF NOT A = 1 DISPLAY "Y" END-IF.
    STOP RUN.

