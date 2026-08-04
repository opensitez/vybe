*> vybe-test: cobol/move_group_redefines/redefines_signed_view_of_unsigned
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BALANCE PIC 9(6) VALUE 100000.
01 SIGNED-BAL REDEFINES BALANCE PIC S9(6).
PROCEDURE DIVISION.
    DISPLAY SIGNED-BAL.
    STOP RUN.

