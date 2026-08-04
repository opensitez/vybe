*> vybe-test: cobol/module_program_linkage/main_program_calls_worker_program_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-PROG.
PROCEDURE DIVISION.
    CALL "SUBPROG1".
    STOP RUN.

