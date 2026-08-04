*> vybe-test: cobol/string_functions_intrinsic/intrinsic_substitute_multi_char_old
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(20) VALUE "HELLO WORLD".
01 R PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(S "WORLD" "COBOL") TO R.
    STOP RUN.

