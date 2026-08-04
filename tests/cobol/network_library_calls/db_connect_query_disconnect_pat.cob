*> vybe-test: cobol/network_library_calls/db_connect_query_disconnect_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_network_library_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN PIC X(100) VALUE "sqlite:test.db".
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL SELECT 1 END-EXEC.
    EXEC SQL COMMIT END-EXEC.
    STOP RUN.

