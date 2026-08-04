*> vybe-test: cobol/strings/func_reverse
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(10) VALUE "Hello".
01 R PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION REVERSE(TXT) TO R.
    DISPLAY R.
    STOP RUN.

