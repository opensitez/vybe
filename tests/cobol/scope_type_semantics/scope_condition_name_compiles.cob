*> vybe-test: cobol/scope_type_semantics/scope_condition_name_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9.
   88 ONN VALUE 1.
PROCEDURE DIVISION.
    SET ONN TO TRUE.
    STOP RUN.

