*> vybe-test: cobol/decimal_point_clause/decimal_point_move_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC4.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9V99 VALUE 1,25.
01 B PIC 9V99.
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

