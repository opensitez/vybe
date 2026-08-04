*> vybe-test: cobol/control_flow_calls_matrix/if_case_03
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    IF A = 1 AND B = 2 DISPLAY "Y" END-IF.
    STOP RUN.

