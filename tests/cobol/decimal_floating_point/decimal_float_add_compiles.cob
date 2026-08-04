*> vybe-test: cobol/decimal_floating_point/decimal_float_add_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DFP5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE FLOAT-DECIMAL-16.
01 B USAGE FLOAT-DECIMAL-16.
PROCEDURE DIVISION.
    ADD B TO A.
    STOP RUN.

