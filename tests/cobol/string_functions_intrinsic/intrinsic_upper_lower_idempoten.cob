*> vybe-test: cobol/string_functions_intrinsic/intrinsic_upper_lower_idempotent
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "HELLO".
01 R PIC X(5).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(FUNCTION LOWER-CASE(S)) TO R.
    STOP RUN.

