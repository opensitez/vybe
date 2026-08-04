*> vybe-test: cobol/module_calls_program_linkage/call_returning_handle_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 H PIC X(20).
PROCEDURE DIVISION.
    CALL "E" RETURNING H.
    STOP RUN.

