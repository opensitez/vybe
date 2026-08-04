*> vybe-test: cobol/builtin_libraries_common/sql_connect_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(100) VALUE "sqlite:test.db".
PROCEDURE DIVISION.
    EXEC SQL CONNECT :D END-EXEC.
    STOP RUN.

