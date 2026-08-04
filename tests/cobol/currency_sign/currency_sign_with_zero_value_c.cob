*> vybe-test: cobol/currency_sign/currency_sign_with_zero_value_compiles
*> origin: languages/cobol/tests/cobol/test_currency_sign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CUR10.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CURRENCY SIGN IS "$".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DST PIC $ZZ9.99 VALUE 0.
PROCEDURE DIVISION.
    STOP RUN.

