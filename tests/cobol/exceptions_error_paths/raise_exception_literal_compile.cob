*> vybe-test: cobol/exceptions_error_paths/raise_exception_literal_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    RAISE EXCEPTION "ERR".
    STOP RUN.

