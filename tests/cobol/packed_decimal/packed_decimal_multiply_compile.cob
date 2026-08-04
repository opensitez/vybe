*> vybe-test: cobol/packed_decimal/packed_decimal_multiply_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3) COMP-3 VALUE 3.
01 B PIC S9(3) COMP-3 VALUE 4.
PROCEDURE DIVISION.
    MULTIPLY A BY B.
    STOP RUN.

