*> vybe-test: cobol/literals_strings_interpolation/display_group_literal_mix_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G PIC X(5) VALUE "GROUP".
PROCEDURE DIVISION.
    DISPLAY "VAL:" G.
    STOP RUN.

