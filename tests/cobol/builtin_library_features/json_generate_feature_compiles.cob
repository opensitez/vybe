*> vybe-test: cobol/builtin_library_features/json_generate_feature_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(10) VALUE "ALICE".
   05 WS-AGE PIC 9(3) VALUE 30.
01 WS-JSON PIC X(200).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-REC.
    STOP RUN.

