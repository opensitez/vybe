*> vybe-test: cobol/literals/numeric_and_alpha_literals_compile
*> origin: languages/cobol/tests/cobol/test_literals.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(5) VALUE 100.
01 WS-TXT PIC X(5) VALUE "A".
PROCEDURE DIVISION.
    MOVE 12345 TO WS-NUM.
    MOVE "HELLO" TO WS-TXT.
    STOP RUN.

