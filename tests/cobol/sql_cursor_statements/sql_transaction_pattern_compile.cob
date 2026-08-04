*> vybe-test: cobol/sql_cursor_statements/sql_transaction_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SQLCODE PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL INSERT INTO T(ID) VALUES(1) END-EXEC.
    IF SQLCODE = 0 EXEC SQL COMMIT END-EXEC ELSE EXEC SQL ROLLBACK END-EXEC END-IF.
    STOP RUN.

