*> vybe-test: cobol/decimal_floating_point/decimal_float_subtract_compiles
*> origin: languages/cobol/tests/cobol/test_decimal_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DFP6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE FLOAT-DECIMAL-34.
01 B USAGE FLOAT-DECIMAL-34.
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    STOP RUN.

