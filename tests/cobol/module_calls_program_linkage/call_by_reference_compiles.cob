*> vybe-test: cobol/module_calls_program_linkage/call_by_reference_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(5).
PROCEDURE DIVISION.
    CALL "M4" USING BY REFERENCE A.
    STOP RUN.

