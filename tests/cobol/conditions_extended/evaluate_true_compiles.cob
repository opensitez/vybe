*> vybe-test: cobol/conditions_extended/evaluate_true_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(2) VALUE 85.
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN WS-A >= 90
            DISPLAY "A"
        WHEN WS-A >= 80
            DISPLAY "B"
        WHEN OTHER
            DISPLAY "F"
    END-EVALUATE.
    STOP RUN.

