*> vybe-test: cobol/decimal_point_clause/decimal_point_subtract_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC6.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9V9 VALUE 3,3.
01 B PIC 9V9 VALUE 1,1.
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    STOP RUN.

