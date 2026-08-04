*> vybe-test: cobol/builtin_libraries_common/sql_cursor_close_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC SQL CLOSE C1 END-EXEC.
    STOP RUN.

