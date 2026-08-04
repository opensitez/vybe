*> vybe-test: cobol/network_library_calls/sql_cursor_lifecycle_compiles
*> origin: languages/cobol/tests/cobol/test_network_library_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ID PIC 9(5).
PROCEDURE DIVISION.
    EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM USERS END-EXEC.
    EXEC SQL OPEN C1 END-EXEC.
    EXEC SQL FETCH C1 INTO :WS-ID END-EXEC.
    EXEC SQL CLOSE C1 END-EXEC.
    STOP RUN.

