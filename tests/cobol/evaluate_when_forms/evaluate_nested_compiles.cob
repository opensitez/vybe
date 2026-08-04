*> vybe-test: cobol/evaluate_when_forms/evaluate_nested_compiles
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
01 Y PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE X
        WHEN 1
            EVALUATE Y
                WHEN 2
                    DISPLAY "1-2"
                WHEN OTHER
                    DISPLAY "1-OTHER"
            END-EVALUATE
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    STOP RUN.

