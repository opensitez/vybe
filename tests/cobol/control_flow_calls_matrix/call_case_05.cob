*> vybe-test: cobol/control_flow_calls_matrix/call_case_05
*> origin: languages/cobol/tests/cobol/test_control_flow_calls_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    CALL "SUB5" USING BY VALUE X.
    STOP RUN.

