*> vybe-test: cobol/move_group_redefines/redefines_alpha_over_numeric
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NUM-BASE PIC 9(4) VALUE 1234.
01 ALPHA-VIEW REDEFINES NUM-BASE PIC X(4).
PROCEDURE DIVISION.
    DISPLAY ALPHA-VIEW.
    STOP RUN.

