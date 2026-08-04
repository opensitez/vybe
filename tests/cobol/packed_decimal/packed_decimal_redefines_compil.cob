*> vybe-test: cobol/packed_decimal/packed_decimal_redefines_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD3.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P1 PIC S9(3) COMP-3 VALUE 123.
01 P1-X REDEFINES P1 PIC X(2).
PROCEDURE DIVISION.
    STOP RUN.

