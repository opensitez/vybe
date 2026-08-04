*> vybe-test: cobol/module_program_linkage/call_using_single_argument_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-ONE-ARG.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-V PIC 9(3) VALUE 7.
PROCEDURE DIVISION.
    CALL "SUBPROG3" USING WS-V.
    STOP RUN.

