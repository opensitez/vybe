*> vybe-test: cobol/program_exit/test_exit_program_verb
*> origin: languages/cobol/tests/cobol/test_program_exit.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    EXIT PROGRAM.
    STOP RUN.

