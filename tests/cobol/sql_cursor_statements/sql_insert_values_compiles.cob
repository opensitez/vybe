*> vybe-test: cobol/sql_cursor_statements/sql_insert_values_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(5) VALUE 1.
01 N PIC X(20) VALUE "A".
PROCEDURE DIVISION.
    EXEC SQL INSERT INTO USERS (ID, NAME) VALUES (:I, :N) END-EXEC.
    STOP RUN.

