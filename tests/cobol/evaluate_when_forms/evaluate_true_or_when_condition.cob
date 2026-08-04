*> vybe-test: cobol/evaluate_when_forms/evaluate_true_or_when_condition
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN N = 5 OR N = 7
            DISPLAY "FIVE OR SEVEN"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "FIVE OR SEVEN" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FIVE OR SEVEN"
        DISPLAY "FAIL: want [FIVE OR SEVEN] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

