*> vybe-test: cobol/module_calls_program_linkage/module_sequence_b_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "INIT".
    CALL "RUN".
    CALL "DONE".
    STOP RUN.

