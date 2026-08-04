*> vybe-test: cobol/module_calls_program_linkage/call_after_if_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF X = 1 CALL "C" END-IF.
    STOP RUN.

