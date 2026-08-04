*> vybe-test: cobol/module_calls_program_linkage/dynamic_call_name_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(10) VALUE "M10".
PROCEDURE DIVISION.
    CALL N.
    STOP RUN.

