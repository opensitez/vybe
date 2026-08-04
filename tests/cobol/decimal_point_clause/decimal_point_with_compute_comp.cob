*> vybe-test: cobol/decimal_point_clause/decimal_point_with_compute_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_point_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DPC2.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9V9 VALUE 1,5.
01 B PIC 9V9 VALUE 2,5.
01 R PIC 9V9.
PROCEDURE DIVISION.
    COMPUTE R = A + B.
    STOP RUN.

