*> vybe-test: cobol/module_program_linkage/module_call_with_using_and_exception_branches_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-LINKAGE-EX.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ID PIC 9(3) VALUE 101.
PROCEDURE DIVISION.
    CALL "SUBPROG4" USING WS-ID
        ON EXCEPTION DISPLAY "FAIL"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

