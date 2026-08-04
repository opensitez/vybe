*> vybe-test: cobol/scope/nested_evaluate_if_scope_compiles
*> origin: languages/cobol/tests/cobol/test_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 V PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE V
        WHEN 2
            IF V > 0
                DISPLAY "POS"
            END-IF
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    STOP RUN.

