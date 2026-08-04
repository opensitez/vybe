*> vybe-test: cobol/scope_type_semantics/scope_occurs_data_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC X(2) OCCURS 3 TIMES.
PROCEDURE DIVISION.
    MOVE "AA" TO T(1).
    STOP RUN.

