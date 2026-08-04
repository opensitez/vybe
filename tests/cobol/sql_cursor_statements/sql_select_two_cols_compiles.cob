*> vybe-test: cobol/sql_cursor_statements/sql_select_two_cols_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(20).
01 B PIC 9(5).
PROCEDURE DIVISION.
    EXEC SQL SELECT NAME, ID INTO :A, :B FROM USERS WHERE ID = 1 END-EXEC.
    STOP RUN.

