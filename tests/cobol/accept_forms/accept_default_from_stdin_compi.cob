*> vybe-test: cobol/accept_forms/accept_default_from_stdin_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 INPUT-LINE PIC X(80).
PROCEDURE DIVISION.
    ACCEPT INPUT-LINE.
    STOP RUN.

