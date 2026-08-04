*> vybe-test: cobol/builtin_library_features/json_generate_with_count_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(8) VALUE "BOB".
01 WS-JSON PIC X(200).
01 WS-LEN PIC 9(5).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-REC COUNT IN WS-LEN.
    STOP RUN.

