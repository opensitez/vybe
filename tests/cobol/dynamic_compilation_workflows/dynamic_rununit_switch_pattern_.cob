*> vybe-test: cobol/dynamic_compilation_workflows/dynamic_rununit_switch_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_dynamic_compilation_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-UNIT PIC X(20) VALUE "U1".
PROCEDURE DIVISION.
    CALL "SET-RUNUNIT" USING WS-UNIT.
    DISPLAY "RUNUNIT-SET".
    STOP RUN.

