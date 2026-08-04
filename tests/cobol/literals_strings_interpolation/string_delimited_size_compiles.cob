*> vybe-test: cobol/literals_strings_interpolation/string_delimited_size_compiles
*> origin: languages/cobol/tests/cobol/test_literals_strings_interpolation.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "AB".
01 B PIC X(5) VALUE "CD".
01 O PIC X(10).
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO O.
    STOP RUN.

