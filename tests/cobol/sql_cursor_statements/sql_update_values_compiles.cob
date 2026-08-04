*> vybe-test: cobol/sql_cursor_statements/sql_update_values_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20) VALUE "B".
PROCEDURE DIVISION.
    EXEC SQL UPDATE USERS SET NAME = :N WHERE ID = 1 END-EXEC.
    STOP RUN.

