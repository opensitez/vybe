*> vybe-test: cobol/string_functions_intrinsic/intrinsic_concatenate_three
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION CONCATENATE("A" "B" "C") TO R.
    STOP RUN.

