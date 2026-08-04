*> vybe-test: cobol/scope_type_semantics/scope_group_data_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G.
   05 A PIC X(3).
   05 B PIC 9(2).
PROCEDURE DIVISION.
    MOVE "ABC" TO A.
    STOP RUN.

