*> vybe-test: cobol/embedded_sql/customer_report
*> origin: languages/cobol/tests/cobol/test_embedded_sql.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CUSTREPORT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN    PIC X(100) VALUE "sqlite:customers.db".
01 WS-ID     PIC 9(10).
01 WS-NAME   PIC X(50).
01 WS-CITY   PIC X(30).
01 WS-BAL    PIC 9(10)V99.
01 WS-TOTAL  PIC 9(12)V99 VALUE 0.
01 WS-COUNT  PIC 9(5) VALUE 0.
01 SQLCODE   PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        DECLARE REPORT-CURSOR CURSOR FOR
        SELECT ID, NAME, CITY, BALANCE
        FROM CUSTOMERS
        ORDER BY NAME
    END-EXEC.
    EXEC SQL OPEN REPORT-CURSOR END-EXEC.
    DISPLAY "Customer Report".
    DISPLAY "========================================".
    PERFORM UNTIL SQLCODE NOT = 0
        EXEC SQL
            FETCH REPORT-CURSOR
            INTO :WS-ID, :WS-NAME, :WS-CITY, :WS-BAL
        END-EXEC
        IF SQLCODE = 0
            DISPLAY WS-ID " " WS-NAME " " WS-CITY " " WS-BAL
            ADD WS-BAL TO WS-TOTAL
            ADD 1 TO WS-COUNT
        END-IF
    END-PERFORM.
    EXEC SQL CLOSE REPORT-CURSOR END-EXEC.
    DISPLAY "========================================".
    DISPLAY "Total Customers: " WS-COUNT.
    DISPLAY "Total Balance:   " WS-TOTAL.
    STOP RUN.

