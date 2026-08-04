*> vybe-test: cobol/builtin_library_features/json_generate_with_nested_group_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(10) VALUE "NINA".
   05 WS-ADDR.
      10 WS-CITY PIC X(10) VALUE "PARIS".
01 WS-JSON PIC X(400).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-REC.
    STOP RUN.

