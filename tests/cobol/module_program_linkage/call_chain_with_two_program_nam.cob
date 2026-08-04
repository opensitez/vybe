*> vybe-test: cobol/module_program_linkage/call_chain_with_two_program_names_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-CHAIN.
PROCEDURE DIVISION.
    CALL "SUB1".
    CALL "SUB2".
    STOP RUN.

