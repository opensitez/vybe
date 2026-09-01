*> vybe-test: cobol/dynamic_compilation_workflows/dynamic_load_and_invoke_pattern_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
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

