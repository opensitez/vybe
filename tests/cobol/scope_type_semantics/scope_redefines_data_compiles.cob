*> vybe-test: cobol/scope_type_semantics/scope_redefines_data_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 B PIC X(10).
01 N REDEFINES B PIC 9(10).
PROCEDURE DIVISION.
    MOVE 1 TO N.
    STOP RUN.

