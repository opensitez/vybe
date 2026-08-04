*> vybe-test: cobol/numeric_functions/intrinsic_numval_converts_string
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "  42.5".
01 R PIC 9(5)V9 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION NUMVAL(S).
    STOP RUN.

