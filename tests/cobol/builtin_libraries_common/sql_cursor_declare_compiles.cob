*> vybe-test: cobol/builtin_libraries_common/sql_cursor_declare_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM USERS END-EXEC.
    STOP RUN.

