*> vybe-test: cobol/scope_type_semantics/type_alpha_picx_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(10).
PROCEDURE DIVISION.
    MOVE "A" TO A.
    STOP RUN.

