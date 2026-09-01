*> vybe-test: cobol/module_program_linkage/module_call_chain_with_three_steps_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-CHAIN-3.
PROCEDURE DIVISION.
    CALL "SUBA".
    CALL "SUBB".
    CALL "SUBC".
    STOP RUN.

