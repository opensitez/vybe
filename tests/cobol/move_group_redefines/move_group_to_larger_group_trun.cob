*> vybe-test: cobol/move_group_redefines/move_group_to_larger_group_truncates
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC.
   05 S1 PIC X(4) VALUE "ABCD".
01 DST.
   05 D1 PIC X(4) VALUE "XXXX".
   05 D2 PIC X(4) VALUE "YYYY".
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    STOP RUN.

