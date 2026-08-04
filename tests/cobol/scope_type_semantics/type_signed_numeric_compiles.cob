*> vybe-test: cobol/scope_type_semantics/type_signed_numeric_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(5).
PROCEDURE DIVISION.
    MOVE -12 TO A.
    STOP RUN.

