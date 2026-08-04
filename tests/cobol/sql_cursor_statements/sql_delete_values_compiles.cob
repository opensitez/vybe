*> vybe-test: cobol/sql_cursor_statements/sql_delete_values_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC SQL DELETE FROM USERS WHERE ID = 1 END-EXEC.
    STOP RUN.

