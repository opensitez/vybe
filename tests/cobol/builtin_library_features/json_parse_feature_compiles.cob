*> vybe-test: cobol/builtin_library_features/json_parse_feature_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-JSON PIC X(200) VALUE '{"name":"BOB"}'.
01 WS-REC.
   05 WS-NAME PIC X(10).
PROCEDURE DIVISION.
    JSON PARSE WS-JSON INTO WS-REC.
    STOP RUN.

