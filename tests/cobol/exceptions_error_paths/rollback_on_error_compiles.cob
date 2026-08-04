*> vybe-test: cobol/exceptions_error_paths/rollback_on_error_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SQLCODE PIC S9(9) VALUE 1.
PROCEDURE DIVISION.
    IF SQLCODE NOT = 0
        EXEC SQL ROLLBACK END-EXEC
    END-IF.
    STOP RUN.

