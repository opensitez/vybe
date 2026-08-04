*> vybe-test: cobol/strings/func_upper
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "hello".
01 R PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(TXT) TO R.
    DISPLAY R.
    STOP RUN.

