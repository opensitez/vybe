*> vybe-test: cobol/packed_decimal/packed_decimal_with_signed_fraction_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD9.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3)V9 COMP-3 VALUE 12.3.
PROCEDURE DIVISION.
    STOP RUN.

