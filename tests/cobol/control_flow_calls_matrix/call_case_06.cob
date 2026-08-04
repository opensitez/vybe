*> vybe-test: cobol/control_flow_calls_matrix/call_case_06
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(2).
PROCEDURE DIVISION.
    CALL "SUB6" RETURNING R.
    STOP RUN.

