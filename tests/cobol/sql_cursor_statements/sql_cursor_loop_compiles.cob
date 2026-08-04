*> vybe-test: cobol/sql_cursor_statements/sql_cursor_loop_compiles
*> origin: languages/cobol/tests/cobol/test_sql_cursor_statements.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(5).
01 SQLCODE PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL OPEN C1 END-EXEC.
    PERFORM UNTIL SQLCODE NOT = 0
        EXEC SQL FETCH C1 INTO :I END-EXEC
    END-PERFORM.
    EXEC SQL CLOSE C1 END-EXEC.
    STOP RUN.

