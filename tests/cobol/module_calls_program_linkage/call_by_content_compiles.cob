*> vybe-test: cobol/module_calls_program_linkage/call_by_content_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(5).
PROCEDURE DIVISION.
    CALL "M5" USING BY CONTENT A.
    STOP RUN.

