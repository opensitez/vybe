*> vybe-test: cobol/accept_forms/exit_in_last_paragraph_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM LAST-PARA.
    STOP RUN.
LAST-PARA.
    DISPLAY "LAST".
    EXIT.
    STOP RUN.

