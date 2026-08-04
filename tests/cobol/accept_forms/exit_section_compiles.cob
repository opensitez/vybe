*> vybe-test: cobol/accept_forms/exit_section_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM MY-SECTION.
    STOP RUN.
MY-SECTION SECTION.
    DISPLAY "IN SECTION".
    EXIT SECTION.
    STOP RUN.

