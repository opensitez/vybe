*> vybe-test: cobol/encoding_international_text/special_names_decimal_point_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3)V99.
PROCEDURE DIVISION.
    MOVE 1 TO X.
    STOP RUN.

