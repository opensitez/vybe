*> vybe-test: cobol/signed_arithmetic/signed_absolute_value_with_function
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC S9(5) VALUE -1234.
01 R PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION ABS(N).
    STOP RUN.

