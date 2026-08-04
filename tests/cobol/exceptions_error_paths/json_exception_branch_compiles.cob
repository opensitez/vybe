*> vybe-test: cobol/exceptions_error_paths/json_exception_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 J PIC X(50).
01 R PIC X(10).
PROCEDURE DIVISION.
    JSON PARSE J INTO R.
    STOP RUN.

