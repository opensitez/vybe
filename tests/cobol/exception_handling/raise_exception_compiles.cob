*> vybe-test: cobol/exception_handling/raise_exception_compiles
*> origin: languages/cobol/tests/cobol/test_exception_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    RAISE EXCEPTION "boom".
    STOP RUN.

