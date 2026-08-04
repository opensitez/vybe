*> vybe-test: cobol/module_calls_program_linkage/module_sequence_a_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "M11".
    CALL "M12".
    CALL "M13".
    STOP RUN.

