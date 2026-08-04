*> vybe-test: cobol/packed_decimal/packed_decimal_compute_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD8.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3) COMP-3 VALUE 10.
01 B PIC S9(3) COMP-3 VALUE 20.
01 C PIC S9(4) COMP-3.
PROCEDURE DIVISION.
    COMPUTE C = A + B.
    STOP RUN.

