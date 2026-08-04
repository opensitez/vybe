*> vybe-test: cobol/sql_cursor_statements/sql_fetch_cursor_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(5).
PROCEDURE DIVISION.
    EXEC SQL FETCH C1 INTO :I END-EXEC.
    STOP RUN.

