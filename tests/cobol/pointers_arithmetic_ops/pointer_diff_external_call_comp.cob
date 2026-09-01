*> vybe-test: cobol/pointers_arithmetic_ops/pointer_diff_external_call_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P1 USAGE POINTER.
01 P2 USAGE POINTER.
01 D PIC 9(5).
PROCEDURE DIVISION.
    CALL "PTR-DIFF" USING P1 P2 D.
    STOP RUN.

