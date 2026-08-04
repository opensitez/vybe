*> vybe-test: cobol/literals_strings_interpolation/literal_string_move_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(10).
PROCEDURE DIVISION.
    MOVE "HELLO" TO A.
    STOP RUN.

