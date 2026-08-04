*> vybe-test: cobol/string_functions_intrinsic/intrinsic_combined_upper_and_reverse
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "hello".
01 R PIC X(5).
PROCEDURE DIVISION.
    MOVE FUNCTION REVERSE(FUNCTION UPPER-CASE(S)) TO R.
    STOP RUN.

