*> vybe-test: cobol/module_program_linkage/call_using_single_argument_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-ONE-ARG.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-V PIC 9(3) VALUE 7.
PROCEDURE DIVISION.
    CALL "SUBPROG3" USING WS-V.
    STOP RUN.

