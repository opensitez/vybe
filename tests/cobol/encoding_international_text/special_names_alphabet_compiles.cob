*> vybe-test: cobol/encoding_international_text/special_names_alphabet_compiles
*> origin: languages/cobol/tests/cobol/test_encoding_international_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET A1 IS STANDARD-1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(5).
PROCEDURE DIVISION.
    DISPLAY X.
    STOP RUN.

