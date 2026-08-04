*> vybe-test: cobol/dynamic_compilation_workflows/dynamic_load_with_returning_handle_compiles
*> origin: languages/cobol/tests/cobol/test_dynamic_compilation_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(200) VALUE "CALL.".
01 WS-HANDLE PIC X(40).
PROCEDURE DIVISION.
    CALL "LOAD-PLUGIN" USING WS-SRC RETURNING WS-HANDLE.
    CALL "RUN-PLUGIN" USING WS-HANDLE.
    STOP RUN.

