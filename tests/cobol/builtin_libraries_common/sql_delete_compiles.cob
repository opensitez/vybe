*> vybe-test: cobol/builtin_libraries_common/sql_delete_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC SQL DELETE FROM USERS WHERE ID = 1 END-EXEC.
    STOP RUN.

