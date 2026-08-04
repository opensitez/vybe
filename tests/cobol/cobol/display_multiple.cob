*> vybe-test: cobol/cobol/display_multiple
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. DISPMUL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10) VALUE "Alice".
01 WS-AGE  PIC 9(3)  VALUE 30.
PROCEDURE DIVISION.
    DISPLAY "Name: " WS-NAME " Age: " WS-AGE.
    STOP RUN.

