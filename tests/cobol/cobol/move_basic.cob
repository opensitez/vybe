*> vybe-test: cobol/cobol/move_basic
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MOV.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(10) VALUE SPACES.
01 WS-B PIC 9(5)  VALUE 0.
PROCEDURE DIVISION.
    MOVE "Hello" TO WS-A.
    MOVE 42 TO WS-B.
    STOP RUN.

