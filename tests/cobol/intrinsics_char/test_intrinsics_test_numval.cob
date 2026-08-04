*> vybe-test: cobol/intrinsics_char/test_intrinsics_test_numval
*> origin: languages/cobol/tests/cobol/test_intrinsics_char.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RES PIC 9(9).
PROCEDURE DIVISION.

    COMPUTE WS-RES = FUNCTION TEST-NUMVAL("123.45").
    COMPUTE WS-RES = FUNCTION TEST-NUMVAL("ABC").
    STOP RUN.

