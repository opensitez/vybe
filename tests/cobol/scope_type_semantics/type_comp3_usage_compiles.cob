*> vybe-test: cobol/scope_type_semantics/type_comp3_usage_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(5) USAGE COMP-3.
PROCEDURE DIVISION.
    ADD 1 TO A.
    STOP RUN.

