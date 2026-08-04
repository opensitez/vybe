*> vybe-test: cobol/binary_numeric_types/binary_comp5_move_compiles
*> origin: languages/cobol/tests/cobol/test_binary_numeric_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. BN10.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(9) COMP-5 VALUE 99.
01 B PIC S9(9) COMP-5.
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

