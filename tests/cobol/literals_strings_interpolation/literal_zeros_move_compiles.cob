*> vybe-test: cobol/literals_strings_interpolation/literal_zeros_move_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(5).
PROCEDURE DIVISION.
    MOVE ZEROS TO A.
    STOP RUN.

