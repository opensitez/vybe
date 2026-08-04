*> vybe-test: cobol/strings/func_length
*> origin: languages/cobol/tests/cobol/test_strings.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(20) VALUE "Hello".
01 L PIC 9(5).
PROCEDURE DIVISION.
    MOVE FUNCTION LENGTH(TXT) TO L.
    DISPLAY L.
    STOP RUN.

