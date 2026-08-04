*> vybe-test: cobol/sql_cursor_statements/sql_select_into_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(50).
PROCEDURE DIVISION.
    EXEC SQL SELECT NAME INTO :N FROM USERS WHERE ID = 1 END-EXEC.
    STOP RUN.

