*> vybe-test: cobol/module_calls_program_linkage/chained_calls_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "A".
    CALL "B".
    STOP RUN.

