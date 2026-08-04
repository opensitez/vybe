*> vybe-test: cobol/module_calls_program_linkage/call_dynamic_name_runtime
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(10) VALUE "NO-MOD".
PROCEDURE DIVISION.
    CALL N
        ON EXCEPTION
            DISPLAY "MISS"
    END-CALL
    STOP RUN.

