*> vybe-test: cobol/builtin_library_features/embedded_sql_insert_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ID PIC 9(10) VALUE 1.
01 WS-NAME PIC X(50) VALUE "A".
PROCEDURE DIVISION.
    EXEC SQL INSERT INTO USERS (ID, NAME) VALUES (:WS-ID, :WS-NAME) END-EXEC.
    STOP RUN.

