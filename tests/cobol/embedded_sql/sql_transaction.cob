*> vybe-test: cobol/embedded_sql/sql_transaction
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

    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        INSERT INTO ACCOUNTS (ID, BALANCE)
        VALUES (:WS-ID, :WS-AMT)
    END-EXEC.
    IF SQLCODE = 0
        EXEC SQL COMMIT END-EXEC
        DISPLAY "Committed"
    ELSE
        EXEC SQL ROLLBACK END-EXEC
        DISPLAY "Rolled back"
    END-IF.
    STOP RUN.

