*> vybe-test: cobol/final_features/special_names_decimal_comma
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SPECNAMES.
ENVIRONMENT DIVISION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AMT PIC 9(5)V99 VALUE 1234.56.
PROCEDURE DIVISION.
    DISPLAY WS-AMT.
    STOP RUN.

