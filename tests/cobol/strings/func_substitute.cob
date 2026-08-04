*> vybe-test: cobol/strings/func_substitute
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(30) VALUE "Hello World".
01 R PIC X(30).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(TXT "World" "COBOL") TO R.
    DISPLAY R.
    STOP RUN.

