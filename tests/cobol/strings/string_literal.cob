*> vybe-test: cobol/strings/string_literal
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC X(30).
PROCEDURE DIVISION.
    STRING "Hello" DELIMITED BY SIZE " World" DELIMITED BY SIZE INTO R.
    STOP RUN.

