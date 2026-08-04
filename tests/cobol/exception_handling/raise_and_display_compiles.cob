*> vybe-test: cobol/exception_handling/raise_and_display_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    RAISE EXCEPTION "fatal".
    DISPLAY "after error".
    STOP RUN.

