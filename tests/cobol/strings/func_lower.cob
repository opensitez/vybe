*> vybe-test: cobol/strings/func_lower
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "HELLO".
01 R PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION LOWER-CASE(TXT) TO R.
    DISPLAY R.
    STOP RUN.

