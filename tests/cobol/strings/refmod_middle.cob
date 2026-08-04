*> vybe-test: cobol/strings/refmod_middle
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "Hello World".
01 SUB PIC X(5).
PROCEDURE DIVISION.
    MOVE TXT(7:5) TO SUB.
    DISPLAY SUB.
    STOP RUN.

