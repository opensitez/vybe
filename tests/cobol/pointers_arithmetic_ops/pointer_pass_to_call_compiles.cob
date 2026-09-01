*> vybe-test: cobol/pointers_arithmetic_ops/pointer_pass_to_call_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    CALL "PTR-USE" USING P.
    STOP RUN.

