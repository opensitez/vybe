*> vybe-test: cobol/international_text_support/special_names_with_alphabet_compiles
*> origin: languages/cobol/tests/cobol/test_international_text_support.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET MY-ALPHA IS STANDARD-1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "ABC".
PROCEDURE DIVISION.
    DISPLAY WS-TXT.
    STOP RUN.

