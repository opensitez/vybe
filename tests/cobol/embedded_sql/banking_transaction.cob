*> vybe-test: cobol/embedded_sql/banking_transaction
*> origin: languages/cobol/tests/cobol/test_embedded_sql.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. BANKING.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN     PIC X(100) VALUE "sqlite:bank.db".
01 WS-FROM-ID PIC 9(10) VALUE 1001.
01 WS-TO-ID   PIC 9(10) VALUE 1002.
01 WS-AMOUNT  PIC 9(10)V99 VALUE 500.00.
01 WS-BAL     PIC 9(10)V99 VALUE 0.
01 SQLCODE    PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        SELECT BALANCE INTO :WS-BAL
        FROM ACCOUNTS
        WHERE ACCOUNT_ID = :WS-FROM-ID
    END-EXEC.
    IF WS-BAL >= WS-AMOUNT
        EXEC SQL
            UPDATE ACCOUNTS
            SET BALANCE = BALANCE - :WS-AMOUNT
            WHERE ACCOUNT_ID = :WS-FROM-ID
        END-EXEC
        EXEC SQL
            UPDATE ACCOUNTS
            SET BALANCE = BALANCE + :WS-AMOUNT
            WHERE ACCOUNT_ID = :WS-TO-ID
        END-EXEC
        EXEC SQL COMMIT END-EXEC
        DISPLAY "Transfer complete"
    ELSE
        DISPLAY "Insufficient funds"
    END-IF.
    STOP RUN.

