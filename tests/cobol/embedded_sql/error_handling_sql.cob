*> vybe-test: cobol/embedded_sql/error_handling_sql
*> origin: languages/cobol/tests/cobol/test_embedded_sql.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SQLERR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN    PIC X(100) VALUE "sqlite:test.db".
01 WS-NAME   PIC X(50).
01 WS-ID     PIC 9(10) VALUE 99999.
01 SQLCODE   PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        SELECT NAME INTO :WS-NAME
        FROM USERS WHERE ID = :WS-ID
    END-EXEC.
    EVALUATE SQLCODE
        WHEN 0
            DISPLAY "Found: " WS-NAME
        WHEN 100
            DISPLAY "No data found"
        WHEN OTHER
            DISPLAY "SQL Error: " SQLCODE
    END-EVALUATE.
    STOP RUN.

