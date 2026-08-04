*> vybe-test: cobol/numeric_functions/intrinsic_substitute_replaces_chars
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "HELLO".
01 R PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(S "L" "R") TO R.
    STOP RUN.

