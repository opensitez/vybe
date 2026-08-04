*> vybe-test: cobol/accept_forms/accept_from_console_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(20).
PROCEDURE DIVISION.
    ACCEPT S FROM CONSOLE.
    STOP RUN.

