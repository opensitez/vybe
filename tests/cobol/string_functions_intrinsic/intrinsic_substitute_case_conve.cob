*> vybe-test: cobol/string_functions_intrinsic/intrinsic_substitute_case_converts
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "Hello".
01 R PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(S "H" "h") TO R.
    STOP RUN.

