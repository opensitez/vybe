*> vybe-test: cobol/numeric_functions/intrinsic_integer_of_date_day_of_integer_roundtrip
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC 9(8) VALUE 20230101.
01 I PIC 9(8) VALUE 0.
01 D2 PIC 9(8) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE I = FUNCTION INTEGER-OF-DATE(D).
    COMPUTE D2 = FUNCTION DATE-OF-INTEGER(I).
    STOP RUN.

