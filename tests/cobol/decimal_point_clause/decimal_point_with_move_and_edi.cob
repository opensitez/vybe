*> vybe-test: cobol/decimal_point_clause/decimal_point_with_move_and_editing_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC10.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC 9V99 VALUE 4,75.
01 DST PIC ZZ9,99.
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    STOP RUN.

