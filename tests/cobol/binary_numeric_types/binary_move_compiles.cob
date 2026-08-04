*> vybe-test: cobol/binary_numeric_types/binary_move_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN8.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) COMP VALUE 10.
01 B PIC S9(4) COMP.
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

