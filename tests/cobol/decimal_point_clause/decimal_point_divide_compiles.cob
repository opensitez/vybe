*> vybe-test: cobol/decimal_point_clause/decimal_point_divide_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC8.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9V9 VALUE 2,0.
01 B PIC 9V9 VALUE 6,0.
PROCEDURE DIVISION.
    DIVIDE A INTO B.
    STOP RUN.

