*> vybe-test: cobol/intrinsics_statistical/test_intrinsics_financial
*> origin: languages/cobol/tests/cobol/test_intrinsics_statistical.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ANN PIC 9(3)V9999.
01 WS-PV PIC 9(5)V99.
PROCEDURE DIVISION.

    COMPUTE WS-ANN = FUNCTION ANNUITY(0.05 10).
    COMPUTE WS-PV = FUNCTION PRESENT-VALUE(0.05 100 100 100).
    STOP RUN.

