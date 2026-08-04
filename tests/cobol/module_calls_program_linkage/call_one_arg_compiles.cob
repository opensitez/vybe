*> vybe-test: cobol/module_calls_program_linkage/call_one_arg_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5).
PROCEDURE DIVISION.
    CALL "M2" USING A.
    STOP RUN.

