*> vybe-test: cobol/packed_decimal/packed_decimal_add_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD2.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3) COMP-3 VALUE 10.
01 B PIC S9(3) COMP-3 VALUE 5.
PROCEDURE DIVISION.
    ADD B TO A.
    STOP RUN.

