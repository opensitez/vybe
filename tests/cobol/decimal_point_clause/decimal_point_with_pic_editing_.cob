*> vybe-test: cobol/decimal_point_clause/decimal_point_with_pic_editing_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC3.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC ZZ9,99 VALUE 123,45.
PROCEDURE DIVISION.
    STOP RUN.

