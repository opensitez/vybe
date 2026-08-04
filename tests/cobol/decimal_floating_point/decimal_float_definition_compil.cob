*> vybe-test: cobol/decimal_floating_point/decimal_float_definition_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DFP1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DF USAGE FLOAT-DECIMAL-16.
PROCEDURE DIVISION.
    STOP RUN.

