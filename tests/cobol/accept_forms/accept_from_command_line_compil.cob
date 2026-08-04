*> vybe-test: cobol/accept_forms/accept_from_command_line_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ARG PIC X(80).
PROCEDURE DIVISION.
    ACCEPT ARG FROM COMMAND-LINE.
    STOP RUN.

