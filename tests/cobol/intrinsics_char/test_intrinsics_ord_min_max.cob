*> vybe-test: cobol/intrinsics_char/test_intrinsics_ord_min_max
*> origin: languages/cobol/tests/cobol/test_intrinsics_char.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ORD PIC 9(5).
PROCEDURE DIVISION.

    COMPUTE WS-ORD = FUNCTION ORD-MAX.
    COMPUTE WS-ORD = FUNCTION ORD-MIN.
    STOP RUN.

