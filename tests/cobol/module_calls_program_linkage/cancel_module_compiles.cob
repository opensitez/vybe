*> vybe-test: cobol/module_calls_program_linkage/cancel_module_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CANCEL "M9".
    STOP RUN.

