*> vybe-test: cobol/builtin_libraries_common/sql_commit_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC SQL COMMIT END-EXEC.
    STOP RUN.

