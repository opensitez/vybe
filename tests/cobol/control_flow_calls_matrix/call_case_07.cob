*> vybe-test: cobol/control_flow_calls_matrix/call_case_07
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PGM PIC X(8) VALUE "SUB7".
PROCEDURE DIVISION.
    CALL PGM.
    STOP RUN.

