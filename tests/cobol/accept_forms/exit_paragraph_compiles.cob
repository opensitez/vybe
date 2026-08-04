*> vybe-test: cobol/accept_forms/exit_paragraph_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM MY-PARA.
    STOP RUN.
MY-PARA.
    DISPLAY "PARA".
    EXIT.
    STOP RUN.

