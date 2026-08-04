*> vybe-test: cobol/strings/func_concatenate
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(10) VALUE "Hello".
01 B PIC X(10) VALUE "World".
01 R PIC X(25).
PROCEDURE DIVISION.
    MOVE FUNCTION CONCATENATE(A B) TO R.
    DISPLAY R.
    STOP RUN.

