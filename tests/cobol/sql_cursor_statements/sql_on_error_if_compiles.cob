*> vybe-test: cobol/sql_cursor_statements/sql_on_error_if_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SQLCODE PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL SELECT 1 END-EXEC.
    IF SQLCODE NOT = 0 DISPLAY "E" END-IF.
    STOP RUN.

