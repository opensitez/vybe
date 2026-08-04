*> vybe-test: cobol/binary_numeric_types/binary_subtraction_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) COMP VALUE 10.
01 B PIC S9(4) COMP VALUE 5.
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    STOP RUN.

