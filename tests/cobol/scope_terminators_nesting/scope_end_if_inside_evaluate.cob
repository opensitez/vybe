*> vybe-test: cobol/scope_terminators_nesting/scope_end_if_inside_evaluate
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 1.
PROCEDURE DIVISION.
    EVALUATE N
        WHEN 1
            IF N = 1
                DISPLAY "ONE"
            END-IF
    END-EVALUATE.
    STOP RUN.

