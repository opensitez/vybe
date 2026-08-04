*> vybe-test: cobol/dynamic_compilation_workflows/dynamic_load_and_invoke_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_dynamic_compilation_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MODULE PIC X(40) VALUE "PLUGIN-A".
PROCEDURE DIVISION.
    CALL "LOAD-MODULE" USING WS-MODULE.
    CALL "INVOKE-MODULE" USING WS-MODULE.
    STOP RUN.

