*> vybe-test: cobol/arithmetic_control_flow_matrix/call_returning_compiles
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(3).
PROCEDURE DIVISION.
    CALL "SUBRET" RETURNING R.
    STOP RUN.

