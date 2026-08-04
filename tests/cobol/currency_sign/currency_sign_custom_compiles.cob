*> vybe-test: cobol/currency_sign/currency_sign_custom_compiles
*> origin: languages/cobol/tests/cobol/test_currency_sign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CUR2.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CURRENCY SIGN IS "E".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 AMT PIC E9.99.
PROCEDURE DIVISION.
    STOP RUN.

