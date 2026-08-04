*> vybe-test: cobol/dynamic_compilation_workflows/dynamic_dispatch_with_exception_branch_compiles
*> origin: languages/cobol/tests/cobol/test_dynamic_compilation_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TARGET PIC X(20) VALUE "HANDLER".
PROCEDURE DIVISION.
    CALL WS-TARGET
        ON EXCEPTION DISPLAY "CALL-FAIL"
        NOT ON EXCEPTION DISPLAY "CALL-OK"
    END-CALL.
    STOP RUN.

