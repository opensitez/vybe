*> vybe-test: cobol/currency_sign/currency_sign_alternate_character_compiles
*> origin: languages/cobol/tests/cobol/test_currency_sign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CUR16.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CURRENCY SIGN IS "!".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 AMT PIC !9.99 VALUE 1234.56.
PROCEDURE DIVISION.
    DISPLAY AMT.
    STOP RUN.

