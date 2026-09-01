*> vybe-test: cobol/interfaces_dynamic_binding/dynamic_invoke_by_name_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE POINTER.
01 N PIC X(10) VALUE "M1".
PROCEDURE DIVISION.
    CALL "INVOKE-NAME" USING O N.
    STOP RUN.

