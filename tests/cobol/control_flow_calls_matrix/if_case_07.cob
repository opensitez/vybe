*> vybe-test: cobol/control_flow_calls_matrix/if_case_07
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(3) VALUE "ABC".
PROCEDURE DIVISION.
    IF X IS ALPHABETIC DISPLAY "A" END-IF.
    STOP RUN.

