*> vybe-test: cobol/evaluate_when_forms/evaluate_not_value_compiles
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 3.
PROCEDURE DIVISION.
    EVALUATE N
        WHEN NOT 1
            DISPLAY "NOT ONE"
        WHEN OTHER
            DISPLAY "ONE"
    END-EVALUATE.
    STOP RUN.

