*> vybe-test: cobol/scope/evaluate_scope_compiles
*> origin: languages/cobol/tests/cobol/test_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VAL PIC 9(1) VALUE 2.
PROCEDURE DIVISION.
    EVALUATE WS-VAL
        WHEN 1
            DISPLAY "one"
        WHEN 2
            DISPLAY "two"
    END-EVALUATE.
    STOP RUN.

