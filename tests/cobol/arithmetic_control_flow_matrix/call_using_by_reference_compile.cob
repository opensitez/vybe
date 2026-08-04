*> vybe-test: cobol/arithmetic_control_flow_matrix/call_using_by_reference_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    CALL "SUBR" USING BY REFERENCE X.
    STOP RUN.

