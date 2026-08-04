*> vybe-test: cobol/module_calls_program_linkage/call_two_args_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5).
01 B PIC 9(3).
PROCEDURE DIVISION.
    CALL "M3" USING A B.
    STOP RUN.

