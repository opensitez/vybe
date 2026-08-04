*> vybe-test: cobol/intrinsics_string/test_intrinsics_numval_cf
*> origin: languages/cobol/tests/cobol/test_intrinsics_string.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(5)V99.
PROCEDURE DIVISION.

    COMPUTE WS-NUM = FUNCTION NUMVAL-C("$12,345.67").
    COMPUTE WS-NUM = FUNCTION NUMVAL-F("123.45").
    STOP RUN.

