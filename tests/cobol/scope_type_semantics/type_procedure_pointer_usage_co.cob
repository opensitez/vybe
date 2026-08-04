*> vybe-test: cobol/scope_type_semantics/type_procedure_pointer_usage_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE PROCEDURE-POINTER.
PROCEDURE DIVISION.
    DISPLAY "P".
    STOP RUN.

