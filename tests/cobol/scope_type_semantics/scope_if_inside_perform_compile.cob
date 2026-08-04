*> vybe-test: cobol/scope_type_semantics/scope_if_inside_perform_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM 2 TIMES
        IF X = 1 DISPLAY "A" END-IF
    END-PERFORM.
    STOP RUN.

