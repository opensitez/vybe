*> vybe-test: cobol/string_functions_intrinsic/intrinsic_trim_trailing_spaces
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "HELLO     ".
01 R PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION TRIM(S TRAILING) TO R.
    STOP RUN.

