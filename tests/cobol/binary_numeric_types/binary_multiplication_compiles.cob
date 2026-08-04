*> vybe-test: cobol/binary_numeric_types/binary_multiplication_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) COMP VALUE 3.
01 B PIC S9(4) COMP VALUE 4.
PROCEDURE DIVISION.
    MULTIPLY A BY B.
    STOP RUN.

