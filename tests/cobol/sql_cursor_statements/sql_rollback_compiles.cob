*> vybe-test: cobol/sql_cursor_statements/sql_rollback_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC SQL ROLLBACK END-EXEC.
    STOP RUN.

