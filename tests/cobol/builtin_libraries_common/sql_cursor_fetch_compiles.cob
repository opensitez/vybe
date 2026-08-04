*> vybe-test: cobol/builtin_libraries_common/sql_cursor_fetch_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(5).
PROCEDURE DIVISION.
    EXEC SQL FETCH C1 INTO :I END-EXEC.
    STOP RUN.

