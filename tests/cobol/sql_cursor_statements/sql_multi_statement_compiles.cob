*> vybe-test: cobol/sql_cursor_statements/sql_multi_statement_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(5) VALUE 1.
PROCEDURE DIVISION.
    EXEC SQL INSERT INTO T(ID) VALUES(:I) END-EXEC.
    EXEC SQL UPDATE T SET ID = :I END-EXEC.
    STOP RUN.

