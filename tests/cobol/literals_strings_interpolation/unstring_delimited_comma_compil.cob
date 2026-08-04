*> vybe-test: cobol/literals_strings_interpolation/unstring_delimited_comma_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "A,B".
01 A PIC X(5).
01 B PIC X(5).
PROCEDURE DIVISION.
    UNSTRING S DELIMITED BY "," INTO A B.
    STOP RUN.

