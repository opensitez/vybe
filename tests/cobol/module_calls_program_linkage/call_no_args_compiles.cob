*> vybe-test: cobol/module_calls_program_linkage/call_no_args_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "M1".
    STOP RUN.

