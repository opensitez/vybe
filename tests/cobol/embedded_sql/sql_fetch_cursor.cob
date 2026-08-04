*> vybe-test: cobol/embedded_sql/sql_fetch_cursor
*> origin: languages/cobol/tests/cobol/test_embedded_sql.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ID    PIC 9(10) VALUE 0.
01 WS-NAME  PIC X(50).
01 WS-AMT   PIC 9(10)V99 VALUE 0.
01 WS-DSN   PIC X(100) VALUE "sqlite:test.db".
01 SQLCODE  PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL FETCH C1 INTO :WS-ID, :WS-NAME END-EXEC.
    STOP RUN.

