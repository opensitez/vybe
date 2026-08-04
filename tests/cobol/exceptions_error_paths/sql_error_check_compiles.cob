*> vybe-test: cobol/exceptions_error_paths/sql_error_check_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SQLCODE PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL SELECT 1 END-EXEC.
    IF SQLCODE NOT = 0 DISPLAY "SQLERR" END-IF.
    STOP RUN.

