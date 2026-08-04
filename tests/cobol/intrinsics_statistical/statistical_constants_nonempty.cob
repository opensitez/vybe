*> vybe-test: cobol/intrinsics_statistical/statistical_constants_nonempty
*> origin: languages/cobol/tests/cobol/test_intrinsics_statistical.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VALUE PIC 9V9 VALUE 0.
PROCEDURE DIVISION.

    COMPUTE WS-VALUE = FUNCTION E + FUNCTION PI.
    COMPUTE WS-VALUE = FUNCTION SIN(0).
    STOP RUN.

