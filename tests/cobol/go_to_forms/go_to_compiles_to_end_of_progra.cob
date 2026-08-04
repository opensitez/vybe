*> vybe-test: cobol/go_to_forms/go_to_compiles_to_end_of_program
*> origin: languages/cobol/tests/cobol/test_go_to_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    GO TO PROGRAM-END.
    DISPLAY "DEAD".
    STOP RUN.
PROGRAM-END.
    DISPLAY "END".
    STOP RUN.

