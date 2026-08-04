*> vybe-test: cobol/builtin_libraries_common/sql_select_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20).
PROCEDURE DIVISION.
    EXEC SQL SELECT NAME INTO :N FROM USERS WHERE ID = 1 END-EXEC.
    STOP RUN.

