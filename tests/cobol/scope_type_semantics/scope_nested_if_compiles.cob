*> vybe-test: cobol/scope_type_semantics/scope_nested_if_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF X = 1
        IF X > 0 DISPLAY "Y" END-IF
    END-IF.
    STOP RUN.

