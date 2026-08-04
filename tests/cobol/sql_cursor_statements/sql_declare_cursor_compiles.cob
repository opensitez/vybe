*> vybe-test: cobol/sql_cursor_statements/sql_declare_cursor_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM USERS END-EXEC.
    STOP RUN.

