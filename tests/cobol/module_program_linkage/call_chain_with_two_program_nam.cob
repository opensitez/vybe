*> vybe-test: cobol/module_program_linkage/call_chain_with_two_program_names_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-CHAIN.
PROCEDURE DIVISION.
    CALL "SUB1".
    CALL "SUB2".
    STOP RUN.

