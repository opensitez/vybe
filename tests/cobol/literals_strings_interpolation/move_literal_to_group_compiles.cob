*> vybe-test: cobol/literals_strings_interpolation/move_literal_to_group_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G.
   05 A PIC X(3).
   05 B PIC X(3).
PROCEDURE DIVISION.
    MOVE "ABCDEF" TO G.
    STOP RUN.

