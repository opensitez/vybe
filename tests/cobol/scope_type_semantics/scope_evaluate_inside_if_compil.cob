*> vybe-test: cobol/scope_type_semantics/scope_evaluate_inside_if_compiles
*> origin: languages/cobol/tests/cobol/test_scope_type_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 2.
PROCEDURE DIVISION.
    IF X > 0
        EVALUATE X
            WHEN 1 DISPLAY "A"
            WHEN OTHER DISPLAY "B"
        END-EVALUATE
    END-IF.
    STOP RUN.

