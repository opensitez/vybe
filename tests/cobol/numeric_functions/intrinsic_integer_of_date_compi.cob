*> vybe-test: cobol/numeric_functions/intrinsic_integer_of_date_compiles
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(8) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION INTEGER-OF-DATE(20230101).
    STOP RUN.

