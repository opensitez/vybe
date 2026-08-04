*> vybe-test: cobol/scope_type_semantics/type_numeric_pic9_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(5).
PROCEDURE DIVISION.
    MOVE 12 TO A.
    STOP RUN.

