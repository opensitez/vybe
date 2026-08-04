*> vybe-test: cobol/numeric_functions/intrinsic_numval_c_compiles
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "$1,234.56".
01 R PIC 9(8)V99 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION NUMVAL-C(S "$").
    STOP RUN.

