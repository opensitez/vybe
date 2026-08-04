*> vybe-test: cobol/currency_sign/currency_sign_with_plus_minus_picture_compiles
*> origin: languages/cobol/tests/cobol/test_currency_sign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CUR8.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CURRENCY SIGN IS "$".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DST PIC +$9.99.
PROCEDURE DIVISION.
    STOP RUN.

