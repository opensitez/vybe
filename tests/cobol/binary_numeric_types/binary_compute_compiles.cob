*> vybe-test: cobol/binary_numeric_types/binary_compute_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN9.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) COMP VALUE 10.
01 B PIC S9(4) COMP VALUE 5.
01 C PIC S9(4) COMP.
PROCEDURE DIVISION.
    COMPUTE C = A + B.
    STOP RUN.

