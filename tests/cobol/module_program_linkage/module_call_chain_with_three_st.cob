*> vybe-test: cobol/module_program_linkage/module_call_chain_with_three_steps_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-CHAIN-3.
PROCEDURE DIVISION.
    CALL "SUBA".
    CALL "SUBB".
    CALL "SUBC".
    STOP RUN.

