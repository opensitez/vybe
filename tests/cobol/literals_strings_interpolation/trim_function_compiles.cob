*> vybe-test: cobol/literals_strings_interpolation/trim_function_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "  HI  ".
01 O PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION TRIM(S) TO O.
    STOP RUN.

