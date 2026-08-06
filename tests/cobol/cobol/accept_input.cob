*> vybe-test: cobol/cobol/accept_input
*> origin: languages/cobol/tests/cobol/test_cobol.rs
*> vybe-test-mode: compile

IDENTIFICATION DIVISION.
PROGRAM-ID. INPUT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20).
PROCEDURE DIVISION.
    DISPLAY "Enter name: ".
    ACCEPT WS-NAME.
    DISPLAY "Hello " WS-NAME.
    STOP RUN.

