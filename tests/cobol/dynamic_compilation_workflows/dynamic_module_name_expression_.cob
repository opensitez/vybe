*> vybe-test: cobol/dynamic_compilation_workflows/dynamic_module_name_expression_compiles
*> origin: languages/cobol/tests/cobol/test_dynamic_compilation_workflows.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PROG PIC X(20) VALUE "HANDLER".
PROCEDURE DIVISION.
    CALL WS-PROG
        ON EXCEPTION DISPLAY "ERR"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

