*> vybe-test: cobol/dynamic_compilation_workflows/dynamic_compile_external_service_call_compiles
*> origin: languages/cobol/tests/cobol/test_dynamic_compilation_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SOURCE PIC X(200).
01 WS-HANDLE PIC X(40).
PROCEDURE DIVISION.
    CALL "DYNAMIC-COMPILE" USING WS-SOURCE RETURNING WS-HANDLE.
    STOP RUN.

