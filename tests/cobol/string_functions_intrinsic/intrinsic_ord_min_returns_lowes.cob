*> vybe-test: cobol/string_functions_intrinsic/intrinsic_ord_min_returns_lowest
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION ORD-MIN("apple" "banana" "cherry").
    STOP RUN.

