*> vybe-test: cobol/module_program_linkage/call_using_multiple_args_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-PROG-ARGS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10) VALUE "ALICE".
01 WS-ID PIC 9(5) VALUE 42.
PROCEDURE DIVISION.
    CALL "SUBPROG2" USING WS-NAME WS-ID.
    STOP RUN.

