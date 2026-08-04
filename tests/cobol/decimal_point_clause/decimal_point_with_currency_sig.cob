*> vybe-test: cobol/decimal_point_clause/decimal_point_with_currency_sign_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC9.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
    CURRENCY SIGN IS "$".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC $9,99 VALUE $1,25.
PROCEDURE DIVISION.
    STOP RUN.

