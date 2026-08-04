*> vybe-test: cobol/special_registers_detail/tally_special_register_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    MOVE 0 TO TALLY.
    STOP RUN.

