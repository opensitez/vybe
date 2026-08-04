*> vybe-test: cobol/decimal_floating_point/decimal_float_divide_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DFP8.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE FLOAT-DECIMAL-16.
01 B USAGE FLOAT-DECIMAL-34.
PROCEDURE DIVISION.
    DIVIDE A INTO B.
    STOP RUN.

