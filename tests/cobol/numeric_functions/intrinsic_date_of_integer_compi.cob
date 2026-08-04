*> vybe-test: cobol/numeric_functions/intrinsic_date_of_integer_compiles
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(8) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE R = FUNCTION DATE-OF-INTEGER(738521).
    STOP RUN.

