*> vybe-test: cobol/packed_decimal/comp3_definition_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P1 PIC S9(5)V99 COMP-3.
PROCEDURE DIVISION.
    STOP RUN.

