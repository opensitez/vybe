*> vybe-test: cobol/decimal_point_clause/decimal_point_is_comma_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC1.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9V99 VALUE 12,34.
PROCEDURE DIVISION.
    STOP RUN.

