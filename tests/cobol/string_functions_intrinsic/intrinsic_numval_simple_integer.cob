*> vybe-test: cobol/string_functions_intrinsic/intrinsic_numval_simple_integer
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "42".
01 R PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION NUMVAL(S).
    STOP RUN.

