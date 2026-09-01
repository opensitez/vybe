*> vybe-test: cobol/module_calls_program_linkage/call_with_group_arg_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G.
   05 A PIC X(3).
   05 B PIC 9(2).
PROCEDURE DIVISION.
    CALL "MG" USING G.
    STOP RUN.

