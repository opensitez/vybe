*> vybe-test: cobol/builtin_library_features/embedded_sql_select_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ID PIC 9(10) VALUE 1.
01 WS-NAME PIC X(50).
PROCEDURE DIVISION.
    EXEC SQL SELECT NAME INTO :WS-NAME FROM USERS WHERE ID = :WS-ID END-EXEC.
    STOP RUN.

