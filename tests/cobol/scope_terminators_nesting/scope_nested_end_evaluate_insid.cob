*> vybe-test: cobol/scope_terminators_nesting/scope_nested_end_evaluate_inside_if
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    IF A > 0
        EVALUATE B
            WHEN 2
                DISPLAY "A-POS-B-2"
        END-EVALUATE
    END-IF.
    STOP RUN.

