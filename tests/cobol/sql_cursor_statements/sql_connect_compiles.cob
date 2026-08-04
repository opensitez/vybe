*> vybe-test: cobol/sql_cursor_statements/sql_connect_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DSN PIC X(100) VALUE "sqlite:test.db".
PROCEDURE DIVISION.
    EXEC SQL CONNECT :DSN END-EXEC.
    STOP RUN.

