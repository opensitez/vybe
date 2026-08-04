*> vybe-test: cobol/packed_decimal/packed_decimal_divide_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3) COMP-3 VALUE 2.
01 B PIC S9(3) COMP-3 VALUE 8.
PROCEDURE DIVISION.
    DIVIDE A INTO B.
    STOP RUN.

