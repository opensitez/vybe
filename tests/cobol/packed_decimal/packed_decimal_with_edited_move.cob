*> vybe-test: cobol/packed_decimal/packed_decimal_with_edited_move_compiles
*> origin: languages/cobol/tests/cobol/test_packed_decimal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PD10.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3) COMP-3 VALUE 123.
01 B PIC ZZZ.
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

