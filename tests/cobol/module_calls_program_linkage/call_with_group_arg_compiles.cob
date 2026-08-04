*> vybe-test: cobol/module_calls_program_linkage/call_with_group_arg_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G.
   05 A PIC X(3).
   05 B PIC 9(2).
PROCEDURE DIVISION.
    CALL "MG" USING G.
    STOP RUN.

