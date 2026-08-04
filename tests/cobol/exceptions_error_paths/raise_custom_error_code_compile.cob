*> vybe-test: cobol/exceptions_error_paths/raise_custom_error_code_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 E PIC 9(4) VALUE 1001.
PROCEDURE DIVISION.
    CALL "RAISE-CODE" USING E.
    STOP RUN.

