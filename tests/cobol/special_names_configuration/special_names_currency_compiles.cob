*> vybe-test: cobol/special_names_configuration/special_names_currency_compiles
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CURRENCY SIGN IS "$".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC $$9.99.
PROCEDURE DIVISION.
    MOVE 1 TO A.
    STOP RUN.

