*> vybe-test: cobol/literals_strings_interpolation/display_literal_concat_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(5) VALUE "A".
PROCEDURE DIVISION.
    DISPLAY "X" N "Y".
    STOP RUN.

