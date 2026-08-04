*> vybe-test: cobol/move_group_redefines/redefines_numeric_comp_over_alpha
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BUF PIC X(4) VALUE "\x00\x00\x00\x01".
01 INT-VIEW REDEFINES BUF PIC 9(9) COMP.
PROCEDURE DIVISION.
    DISPLAY INT-VIEW.
    STOP RUN.

