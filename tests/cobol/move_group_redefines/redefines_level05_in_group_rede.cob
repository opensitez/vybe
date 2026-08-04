*> vybe-test: cobol/move_group_redefines/redefines_level05_in_group_redefine
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RECORD-AREA.
   05 AREA-DATA PIC X(10).
   05 NUMERIC-DATA REDEFINES AREA-DATA PIC 9(10).
PROCEDURE DIVISION.
    MOVE 1234567890 TO NUMERIC-DATA.
    DISPLAY AREA-DATA.
    STOP RUN.

